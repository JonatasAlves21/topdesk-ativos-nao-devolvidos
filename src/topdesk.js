// Chamadas à API do TOPdesk: ativos, histórico, pessoas e chamados.

const {
  TOPDESK_URL, TOPDESK_EMAIL, TOPDESK_TOKEN, STATUS_NAO_DEV, AUTOMATION_CALLER_ID,
} = require("./config");
const { fetchWithRetry } = require("./util");
const { applyHistoryItems, matchIncidents } = require("./regras");

function authHeader() {
  const creds = Buffer.from(`${TOPDESK_EMAIL}:${TOPDESK_TOKEN}`).toString("base64");
  return { Authorization: `Basic ${creds}`, Accept: "application/json" };
}

async function apiFetch(endpoint, params = {}) {
  const url = new URL(`${TOPDESK_URL}${endpoint}`);
  for (const [k, v] of Object.entries(params)) url.searchParams.set(k, String(v));
  const res = await fetch(url.toString(), { headers: authHeader() });
  if (!res.ok && res.status !== 206) throw new Error(`HTTP ${res.status} ${url}`);
  return res.json();
}

async function fetchAssets(templateId, tipo) {
  // Pedido SÓ os campos usados em buildRows. data-de-automacao e status (campos antigos)
  // foram removidos: o status já é filtrado no parâmetro `status=` e data-de-automacao
  // nunca era lido. data-nao-devolvido também foi removido (vem sempre undefined).
  const fields =
    tipo === "Notebook"
      ? "numero-de-patrimonio,modelo,nome-da-maquina,dsp,departamento,centro-de-custo"
      : "numero-de-patrimonio,modelo,nome-do-dispositivo,usuario,departamento,centro-de-custo";

  const result = [];
  let start = 0;

  while (true) {
    const data = await apiFetch("/tas/api/assetmgmt/assets", {
      templateId, pageSize: 100, start, fields, status: STATUS_NAO_DEV, archived: false,
    });
    const batch = data.dataSet || [];
    for (const asset of batch) {
      asset._tipo = tipo;
      result.push(asset);
    }
    if (batch.length < 100) break;
    start += 100;
  }

  return result;
}

// Diagnóstico — coleta tempo/páginas por ativo
const historyStats = [];

// Limite duro de páginas. O endpoint /history/pastItems usa cursor-only com 10 itens/página
// e não filtra por campo, então para ativos sem CC histórico o loop varre TUDO até o fim
// (medido: 1.500-2.300 páginas em alguns tablets, ~5min/ativo). 100 páginas = ~1.000 eventos,
// suficiente para ativos com CC em página razoável; o resto vira "(Não informado)" e cai no
// relatório sem CC.
const MAX_HISTORY_PAGES = 100;

// Busca no histórico do ativo a data da mudança para "Não devolvido" e o último
// centro de custo (só busca o CC se knownCC vier vazio).
async function getHistoryData(assetId, assetName = "", knownCC = null) {
  let histPath = `/tas/api/assetmgmt/assets/${assetId}/history/pastItems`;
  const found = { statusChangeDate: null, lastCentroCusto: knownCC ? String(knownCC) : null };
  const ccFromAsset = !!knownCC; // pulou a busca de CC porque o asset já tinha
  let pages = 0;
  let items_seen = 0;
  const t0 = Date.now();

  while (histPath && pages < MAX_HISTORY_PAGES) {
    const res = await fetchWithRetry(`${TOPDESK_URL}${histPath}`, { headers: authHeader() });
    if (!res.ok) break;
    const data  = await res.json();
    const items = data.items || [];
    pages++;
    items_seen += items.length;

    if (applyHistoryItems(items, found, STATUS_NAO_DEV)) break;
    histPath = data.next ? data.next.replace(TOPDESK_URL, "") : null;
  }

  const { statusChangeDate, lastCentroCusto } = found;
  const hitMaxPages = pages >= MAX_HISTORY_PAGES && (!statusChangeDate || !lastCentroCusto);
  const exhausted   = !histPath && !lastCentroCusto && !ccFromAsset; // varreu até o fim, sem CC

  historyStats.push({
    assetId,
    assetName,
    pages,
    items_seen,
    ms: Date.now() - t0,
    found_status: !!statusChangeDate,
    found_cc: !!lastCentroCusto,
    cc_from_asset: ccFromAsset,
    hit_max_pages: hitMaxPages,
    exhausted,
  });

  return { statusChangeDate, lastCentroCusto };
}

// Recupera a UF a partir do usuário antigo vinculado:
//   - busca a pessoa no Topdesk (por e-mail se contém "@", senão por nome)
//   - retorna o nome da filial (branch.name), que nesta instância É a sigla da UF
const ufCache = new Map();
async function getUserUf(user) {
  const key = (user || "").trim();
  if (!key) return "";
  if (ufCache.has(key)) return ufCache.get(key);

  const isEmail = key.includes("@");
  const field   = isEmail ? "email" : "dynamicName";
  const value   = isEmail ? key : `"${key}"`; // nomes precisam de aspas no FIQL
  const url     = `${TOPDESK_URL}/tas/api/persons?query=${field}==${encodeURIComponent(value)}`;

  let uf = "";
  try {
    const res = await fetchWithRetry(url, { headers: authHeader() });
    if (res.ok) {
      const people = await res.json();
      const branch = Array.isArray(people) && people[0]?.branch?.name;
      if (branch) uf = String(branch).trim().toUpperCase();
    }
  } catch (_) { /* falha de lookup → UF fica vazia */ }

  ufCache.set(key, uf);
  return uf;
}

// Carrega os chamados abertos pela automação e vincula cada um ao seu dispositivo.
async function buildIncidentMap(deviceNames) {
  console.log("  [info] buscando chamados de automação pelo caller...");
  const autoIncidents = [];
  let start = 0;

  while (true) {
    const data = await apiFetch("/tas/api/incidents", {
      pageSize: 100, start, caller: AUTOMATION_CALLER_ID,
    });
    const batch = Array.isArray(data) ? data : [];
    autoIncidents.push(...batch.map((i) => ({
      number:       i.number,
      creationDate: i.creationDate,
      request:      i.request || "",
    })));
    if (batch.length < 100) break;
    start += 100;
  }

  console.log(`  [info] ${autoIncidents.length} chamados da automação carregados`);
  return matchIncidents(autoIncidents, deviceNames);
}

module.exports = {
  MAX_HISTORY_PAGES,
  historyStats,
  fetchAssets,
  getHistoryData,
  getUserUf,
  buildIncidentMap,
};
