require("dotenv").config();
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");

const TOPDESK_URL   = (process.env.TOPDESK_URL || "").replace(/\/$/, "");
const TOPDESK_EMAIL = process.env.TOPDESK_EMAIL || "";
const TOPDESK_TOKEN = process.env.TOPDESK_TOKEN || "";

// SharePoint / Azure AD
const AZURE_TENANT_ID     = process.env.AZURE_TENANT_ID || "";
const AZURE_CLIENT_ID     = process.env.AZURE_CLIENT_ID || "";
const AZURE_CLIENT_SECRET = process.env.AZURE_CLIENT_SECRET || "";
const SHAREPOINT_SITE_URL = (process.env.SHAREPOINT_SITE_URL || "").replace(/\/$/, "");
const SHAREPOINT_FOLDER   = (process.env.SHAREPOINT_FOLDER || "").replace(/^\/|\/$/g, "");
const SHAREPOINT_FILENAME = process.env.SHAREPOINT_FILENAME || "Nao Devolvidos.xlsx";

// Relatório separado dos ativos sem CC (não recuperado nem no asset, nem no histórico).
// Vai para um site/pasta próprio para que a equipe possa preencher manualmente.
// Sem SHAREPOINT_EXAURIDOS_SITE_URL o arquivo é gerado só localmente.
const SHAREPOINT_EXAURIDOS_SITE_URL = (process.env.SHAREPOINT_EXAURIDOS_SITE_URL || "").replace(/\/$/, "");
const SHAREPOINT_EXAURIDOS_FOLDER   = (process.env.SHAREPOINT_EXAURIDOS_FOLDER || "").replace(/^\/|\/$/g, "");
const SHAREPOINT_EXAURIDOS_FILENAME = process.env.SHAREPOINT_EXAURIDOS_FILENAME || "Nao_Devolvidos_Sem_CC.xlsx";

// IDs da instância TOPdesk (templates de ativo, status e pessoa usada pela automação
// que abre os chamados). Variam por instância, então vêm do ambiente.
const TEMPLATE_NOTEBOOK    = process.env.TOPDESK_TEMPLATE_NOTEBOOK || "";
const TEMPLATE_CELULAR     = process.env.TOPDESK_TEMPLATE_CELULAR || "";
const STATUS_NAO_DEV       = process.env.TOPDESK_STATUS_NAO_DEVOLVIDO || "";
const AUTOMATION_CALLER_ID = process.env.TOPDESK_AUTOMATION_CALLER_ID || "";

// Sem estes valores as consultas saem com filtro vazio. O script para antes de
// chamar a API, em vez de varrer ativos e histórico com parâmetros errados.
const REQUIRED_ENV = {
  TOPDESK_URL, TOPDESK_EMAIL, TOPDESK_TOKEN,
  TOPDESK_TEMPLATE_NOTEBOOK: TEMPLATE_NOTEBOOK,
  TOPDESK_TEMPLATE_CELULAR: TEMPLATE_CELULAR,
  TOPDESK_STATUS_NAO_DEVOLVIDO: STATUS_NAO_DEV,
  TOPDESK_AUTOMATION_CALLER_ID: AUTOMATION_CALLER_ID,
};

const COLUMNS = [
  { header: "NOME DO DISPOSITIVO",            key: "nome_dispositivo" },
  { header: "MODELO",                         key: "modelo" },
  { header: "TIPO",                           key: "tipo" },
  { header: "PATRIMÔNIO",                     key: "patrimonio" },
  { header: "USUÁRIO ANTIGO",                 key: "usuario" },
  { header: "CENTRO DE CUSTO",                key: "centro_custo" },
  { header: "UF",                             key: "uf" },
  { header: "DEPARTAMENTO",                   key: "departamento" },
  { header: "DIAS SEM RETORNO",               key: "dias_sem_retorno" },
  { header: "DATA MUDANÇA P/ NÃO DEVOLVIDO",  key: "data_mudanca" },
  { header: "DATA DE NÃO DEVOLVIDO",          key: "data_nao_devolvido" },
  { header: "N° DO CHAMADO DA AUTOMAÇÃO",     key: "num_chamado" },
];

// ── Paleta do relatório ───────────────────────────────────────────────────────
const PAL_DARK   = "FF333333"; // cinza escuro — cabeçalhos
const PAL_ORANGE = "FFF57C20"; // laranja de destaque
const PAL_ALT    = "FFF5F5F5"; // cinza claríssimo — linhas alternadas
const PAL_WHITE  = "FFFFFFFF";
const PAL_TEXT   = "FF1A1A1A"; // texto de dados

// Aging (faixa de prazo sem retorno)
const AGING = [
  { max: 30,  bg: "FFE2EFDA", text: "FF375623", label: "≤ 30 dias  (Em dia)"         },
  { max: 90,  bg: "FFFFEB9C", text: "FF9C5700", label: "31 a 90 dias  (Atenção)"     },
  { max: Infinity, bg: "FFFFC7CE", text: "FF9C0006", label: "> 90 dias  (Crítico)"   },
];

function agingBucket(dias) {
  if (typeof dias !== "number") return null;
  return AGING.find((a) => dias <= a.max) || AGING[AGING.length - 1];
}

// ─────────────────────────────────────────────────────────────────────────────

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
  // Pedido SÓ os campos usados em rowsRaw. data-de-automacao e status (campos antigos)
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

// fetch com retry + backoff — tolera falhas de rede transitórias ("fetch failed")
async function fetchWithRetry(url, options, retries = 4) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      return await fetch(url, options);
    } catch (err) {
      if (attempt === retries) throw err;
      const wait = 500 * Math.pow(2, attempt); // 500ms, 1s, 2s, 4s
      await new Promise((r) => setTimeout(r, wait));
    }
  }
}

// Diagnóstico — coleta tempo/páginas por ativo
const _historyStats = [];

// Limite duro de páginas. O endpoint /history/pastItems usa cursor-only com 10 itens/página
// e não filtra por campo, então para ativos sem CC histórico o loop varre TUDO até o fim
// (medido: 1.500-2.300 páginas em alguns tablets, ~5min/ativo). 100 páginas = ~1.000 eventos,
// suficiente para ativos com CC em página razoável; o resto vira "(Não informado)" e cai no
// relatório de exauridos.
const MAX_HISTORY_PAGES = 100;

// Busca no histórico do ativo:
//   - data em que o status mudou para "Não devolvido"
//   - último centro de custo registrado (só busca se knownCC vier vazio)
async function getHistoryData(assetId, assetName = "", knownCC = null) {
  let histPath = `/tas/api/assetmgmt/assets/${assetId}/history/pastItems`;
  let statusChangeDate = null;
  let lastCentroCusto  = knownCC ? String(knownCC) : null;
  const ccFromAsset    = !!knownCC; // pulou a busca de CC porque o asset já tinha
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

    for (const item of items) {
      const fn = item.field?.fieldName;

      if (!statusChangeDate && fn === "status" && item.field?.newValue === STATUS_NAO_DEV) {
        statusChangeDate = item.date;
      }

      if (!lastCentroCusto && fn === "centro-de-custo") {
        // A automação zera o campo ao marcar o ativo como "Não devolvido";
        // o oldValue preserva o último valor real.
        const val = item.field?.newValue || item.field?.oldValue;
        if (val) lastCentroCusto = String(val);
      }

      if (statusChangeDate && lastCentroCusto) break;
    }

    if (statusChangeDate && lastCentroCusto) break;
    histPath = data.next ? data.next.replace(TOPDESK_URL, "") : null;
  }

  const hitMaxPages = pages >= MAX_HISTORY_PAGES && (!statusChangeDate || !lastCentroCusto);
  const exhausted   = !histPath && !lastCentroCusto && !ccFromAsset; // varreu até o fim, sem CC

  _historyStats.push({
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
const _ufCache = new Map();
async function getUserUf(user) {
  const key = (user || "").trim();
  if (!key) return "";
  if (_ufCache.has(key)) return _ufCache.get(key);

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

  _ufCache.set(key, uf);
  return uf;
}

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
  autoIncidents.sort((a, b) => new Date(a.creationDate) - new Date(b.creationDate));

  const incidentMap = {};
  for (const incident of autoIncidents) {
    const body = incident.request || "";
    for (const name of deviceNames) {
      if (!incidentMap[name] && body.includes(name)) {
        incidentMap[name] = { number: incident.number, date: incident.creationDate };
      }
    }
  }

  return incidentMap;
}

function formatDate(raw) {
  if (!raw) return "";
  try {
    const d = new Date(raw);
    const p = (n) => String(n).padStart(2, "0");
    return `${p(d.getDate())}/${p(d.getMonth() + 1)}/${d.getFullYear()} ${p(d.getHours())}:${p(d.getMinutes())}`;
  } catch (_) {
    return raw;
  }
}

// ── Estilos ───────────────────────────────────────────────────────────────────

function applySheetStyle(ws, rows) {
  ws.columns = COLUMNS.map((c) => ({ header: c.header, key: c.key, width: 20 }));

  const diasColIdx = COLUMNS.findIndex((c) => c.key === "dias_sem_retorno") + 1; // 1-based

  // Cabeçalho
  const headerRow = ws.getRow(1);
  headerRow.eachCell((cell) => {
    cell.font      = { bold: true, color: { argb: PAL_WHITE }, name: "Calibri", size: 10 };
    cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: PAL_DARK } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border    = { bottom: { style: "medium", color: { argb: PAL_ORANGE } } };
  });
  headerRow.height = 32;

  // Linhas de dados
  rows.forEach((row, i) => {
    const excelRow = ws.addRow(row);
    const rowBg    = i % 2 === 1 ? PAL_ALT : PAL_WHITE;

    excelRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.font      = { color: { argb: PAL_TEXT }, name: "Calibri", size: 10 };
      cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: rowBg } };
      cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      cell.border    = { bottom: { style: "thin", color: { argb: "FFE0E0E0" } } };

      // Colorir célula de "DIAS SEM RETORNO" conforme faixa
      if (colNumber === diasColIdx) {
        const bucket = agingBucket(cell.value);
        if (bucket) {
          cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: bucket.bg } };
          cell.font      = { bold: true, color: { argb: bucket.text }, name: "Calibri", size: 10 };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }
    });
    excelRow.height = 22;
  });

  // Largura automática
  ws.columns.forEach((col) => {
    let maxLen = col.header ? col.header.length : 10;
    col.eachCell({ includeEmpty: false }, (cell) => {
      const val = cell.value != null ? String(cell.value) : "";
      if (val.length > maxLen) maxLen = val.length;
    });
    col.width = Math.min(maxLen + 4, 50);
  });

  // Auto-filtro + cabeçalho congelado
  ws.autoFilter = { from: "A1", to: { row: 1, column: COLUMNS.length } };
  ws.views = [{ state: "frozen", ySplit: 1 }];
}

// ── Aba de Resumo ─────────────────────────────────────────────────────────────

function buildSummarySheet(wb, rows) {
  const ws = wb.addWorksheet("Resumo");

  // helpers
  const gc = (addr) => ws.getCell(addr);
  const gr = (n)    => ws.getRow(n);

  function styleCell(addr, { value, bold, size, color, bg, halign, valign, wrapText, borderBottom, borderLeft } = {}) {
    const c = gc(addr);
    if (value !== undefined) c.value = value;
    c.font      = { bold: !!bold, size: size || 10, color: { argb: color || PAL_TEXT }, name: "Calibri" };
    c.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: bg || PAL_WHITE } };
    c.alignment = { horizontal: halign || "left", vertical: valign || "middle", wrapText: !!wrapText };
    const border = {};
    if (borderBottom) border.bottom = { style: borderBottom, color: { argb: PAL_ORANGE } };
    if (borderLeft)   border.left   = { style: borderLeft,   color: { argb: PAL_ORANGE } };
    c.border = border;
  }

  function sectionTitle(addr, text, mergeEnd) {
    if (mergeEnd) ws.mergeCells(`${addr}:${mergeEnd}`);
    styleCell(addr, {
      value: `  ${text}`, bold: true, size: 11,
      color: PAL_WHITE, bg: PAL_DARK, halign: "left",
      borderBottom: "medium", borderLeft: "medium",
    });
  }

  function tableHeader(rowN, cols) {
    cols.forEach((text, i) => {
      const c = gr(rowN).getCell(i + 1);
      c.value     = text;
      c.font      = { bold: true, color: { argb: PAL_WHITE }, name: "Calibri", size: 10 };
      c.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: PAL_DARK } };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border    = { bottom: { style: "medium", color: { argb: PAL_ORANGE } } };
    });
    gr(rowN).height = 24;
  }

  function tableRow(rowN, cols, isAlt, overrideBg, overrideText) {
    cols.forEach((val, i) => {
      const c = gr(rowN).getCell(i + 1);
      c.value     = val;
      c.font      = { color: { argb: overrideText || PAL_TEXT }, name: "Calibri", size: 10 };
      c.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: overrideBg || (isAlt ? PAL_ALT : PAL_WHITE) } };
      c.alignment = { horizontal: i === 0 ? "left" : "center", vertical: "middle" };
      c.border    = { bottom: { style: "thin", color: { argb: "FFE0E0E0" } } };
    });
    gr(rowN).height = 20;
  }

  // ── Colunas ──────────────────────────────────────────────────────────────────
  [26, 13, 13, 10, 10, 12].forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── Título ───────────────────────────────────────────────────────────────────
  ws.mergeCells("A1:F1");
  styleCell("A1", {
    value: "NÃO DEVOLVIDOS — RESUMO EXECUTIVO",
    bold: true, size: 14, color: PAL_WHITE, bg: PAL_DARK, halign: "center",
    borderBottom: "thick",
  });
  gr(1).height = 44;

  ws.mergeCells("A2:F2");
  styleCell("A2", {
    value: `Atualizado em ${new Date().toLocaleDateString("pt-BR")}`,
    size: 9, color: "FF999999", bg: PAL_WHITE, halign: "center",
  });
  gr(2).height = 18;

  gr(3).height = 8; // espaçador

  // ── KPI boxes ────────────────────────────────────────────────────────────────
  const isTablet  = (r) => r.tipo === "Celular" && (r.nome_dispositivo || "").toLowerCase().includes("tablet");
  const total     = rows.length;
  const notebooks = rows.filter((r) => r.tipo === "Notebook").length;
  const tablets   = rows.filter(isTablet).length;
  const celulares = rows.filter((r) => r.tipo === "Celular" && !isTablet(r)).length;

  // Labels (row 4)
  ws.mergeCells("A4:B4"); styleCell("A4", { value: "TOTAL GERAL",         bold: true, size: 9, color: "FF666666", bg: PAL_ALT, halign: "center" });
  ws.mergeCells("C4:D4"); styleCell("C4", { value: "NOTEBOOKS",           bold: true, size: 9, color: "FF666666", bg: PAL_ALT, halign: "center" });
  ws.mergeCells("E4:F4"); styleCell("E4", { value: "CELULARES / TABLETS", bold: true, size: 9, color: "FF666666", bg: PAL_ALT, halign: "center" });
  gr(4).height = 22;

  // Values (row 5)
  ws.mergeCells("A5:B5"); styleCell("A5", { value: total,               bold: true, size: 22, color: PAL_ORANGE, bg: PAL_WHITE, halign: "center" });
  ws.mergeCells("C5:D5"); styleCell("C5", { value: notebooks,           bold: true, size: 22, color: PAL_DARK,   bg: PAL_WHITE, halign: "center" });
  ws.mergeCells("E5:F5"); styleCell("E5", { value: `${celulares} / ${tablets}`, bold: true, size: 22, color: PAL_DARK, bg: PAL_WHITE, halign: "center" });
  gr(5).height = 48;

  // Bordas inferiores dos KPIs
  ["A4","C4","E4"].forEach((a) => { gc(a).border = { ...gc(a).border, top: { style: "thin", color: { argb: "FFD0D0D0" } }, left: { style: "thin", color: { argb: "FFD0D0D0" } }, right: { style: "thin", color: { argb: "FFD0D0D0" } } }; });
  ["A5","C5","E5"].forEach((a) => { gc(a).border = { bottom: { style: "thin", color: { argb: "FFD0D0D0" } }, left: { style: "thin", color: { argb: "FFD0D0D0" } }, right: { style: "thin", color: { argb: "FFD0D0D0" } } }; });

  gr(6).height = 10; // espaçador

  // ── Prazo sem retorno ────────────────────────────────────────────────────────
  const pct = (n) => total > 0 ? `${Math.round((n / total) * 100)}%` : "0%";

  const agingStats = AGING.map((a) => ({
    label: a.label,
    count: rows.filter((r) => {
      if (typeof r.dias_sem_retorno !== "number") return false;
      const prev = AGING[AGING.indexOf(a) - 1];
      const min  = prev ? prev.max : -1;
      return r.dias_sem_retorno > min && r.dias_sem_retorno <= a.max;
    }).length,
    bg:   a.bg,
    text: a.text,
  }));
  const semData = rows.filter((r) => typeof r.dias_sem_retorno !== "number").length;

  sectionTitle("A7", "PRAZO SEM RETORNO", "F7");
  gr(7).height = 28;
  tableHeader(8, ["FAIXA", "QUANTIDADE", "% DO TOTAL"]);

  agingStats.forEach(({ label, count, bg, text }, i) => {
    tableRow(9 + i, [label, count, pct(count)], false, bg, text);
  });
  if (semData > 0) {
    const r = 9 + agingStats.length;
    tableRow(r, ["Sem data registrada", semData, pct(semData)], false, PAL_ALT, PAL_TEXT);
  }

  let nextRow = 9 + agingStats.length + (semData > 0 ? 1 : 0) + 2;

  // ── Por UF ───────────────────────────────────────────────────────────────────
  const byUF = {};
  for (const r of rows) byUF[r.uf || "Outros"] = (byUF[r.uf || "Outros"] || 0) + 1;
  const ufEntries = Object.entries(byUF).sort((a, b) => b[1] - a[1]);

  sectionTitle(`A${nextRow}`, "POR ESTADO (UF)", `C${nextRow}`);
  gr(nextRow).height = 28;
  tableHeader(nextRow + 1, ["UF", "QUANTIDADE", "% DO TOTAL"]);
  ufEntries.forEach(([uf, count], i) => {
    tableRow(nextRow + 2 + i, [uf, count, pct(count)], i % 2 === 1);
  });
  nextRow += 2 + ufEntries.length + 2;

  // ── Por Centro de Custo × Tipo de device ─────────────────────────────────────
  // Pivot: CC → { notebook, celular, tablet, total }
  const byCC = {};
  for (const r of rows) {
    const cc  = r.centro_custo || "(Não informado)";
    if (!byCC[cc]) byCC[cc] = { notebook: 0, celular: 0, tablet: 0 };
    if (r.tipo === "Notebook")                          byCC[cc].notebook++;
    else if (isTablet(r))                               byCC[cc].tablet++;
    else                                                byCC[cc].celular++;
  }
  const ccEntries = Object.entries(byCC)
    .map(([cc, v]) => [cc, v.notebook, v.celular, v.tablet, v.notebook + v.celular + v.tablet])
    .sort((a, b) => b[4] - a[4]); // maior total primeiro

  sectionTitle(`A${nextRow}`, "POR CENTRO DE CUSTO", `F${nextRow}`);
  gr(nextRow).height = 28;
  tableHeader(nextRow + 1, ["CENTRO DE CUSTO", "NOTEBOOKS", "CELULARES", "TABLETS", "TOTAL", "% DO TOTAL"]);
  ccEntries.forEach(([cc, nb, cel, tab, tot], i) => {
    tableRow(nextRow + 2 + i, [cc, nb || "-", cel || "-", tab || "-", tot, pct(tot)], i % 2 === 1);
  });

  // ── Nota de escopo ───────────────────────────────────────────────────────────
  // Evita "chamado de divergência": este consolidado e os arquivos diários por
  // obra (fluxo Power Automate) têm regras diferentes de propósito.
  const notaRow = nextRow + 2 + ccEntries.length + 1;
  ws.mergeCells(`A${notaRow}:F${notaRow}`);
  styleCell(`A${notaRow}`, {
    value:
      "Nota de escopo: este consolidado inclui ativos SEM centro de custo (detalhados em Nao_Devolvidos_Sem_CC.xlsx) e considera " +
      "mudanças de status desde 01/2026. Os arquivos diários ATIVOS_NAO_DEVOLVIDOS_OBRA_XXXX (por obra) omitem ativos sem CC e " +
      "não filtram por data — por isso os totais podem diferir entre os relatórios.",
    size: 8, color: "FF999999", halign: "left", wrapText: true,
  });
  gr(notaRow).height = 34;

  ws.views = [{ state: "frozen", ySplit: 1 }];
}

// ── SharePoint ────────────────────────────────────────────────────────────────

async function getGraphToken() {
  const url = `https://login.microsoftonline.com/${AZURE_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    "client_credentials",
    client_id:     AZURE_CLIENT_ID,
    client_secret: AZURE_CLIENT_SECRET,
    scope:         "https://graph.microsoft.com/.default",
  });
  const res = await fetch(url, { method: "POST", body });
  if (!res.ok) throw new Error(`Erro ao obter token Azure: ${await res.text()}`);
  return (await res.json()).access_token;
}

async function uploadToSharePoint(localFilePath, opts = {}) {
  const siteUrlStr = opts.siteUrl  || SHAREPOINT_SITE_URL;
  const folder     = (opts.folder  ?? SHAREPOINT_FOLDER) || "";
  const filename   = opts.filename || SHAREPOINT_FILENAME;
  const label      = opts.label    || "SharePoint";

  const configured = AZURE_TENANT_ID && AZURE_CLIENT_ID && AZURE_CLIENT_SECRET && siteUrlStr;
  if (!configured) {
    console.log(`  [${label}] Variáveis não configuradas — upload ignorado.`);
    return;
  }

  console.log(`Enviando para ${label}...`);

  const token   = await getGraphToken();
  const headers = { Authorization: `Bearer ${token}` };

  const siteUrl = new URL(siteUrlStr);
  const siteRes = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteUrl.hostname}:${siteUrl.pathname}`,
    { headers }
  );
  if (!siteRes.ok) throw new Error(`Site ${label} não encontrado: ${await siteRes.text()}`);
  const { id: siteId } = await siteRes.json();

  const remotePath  = folder ? `${folder}/${filename}` : filename;
  const encodedPath = remotePath.split("/").map(encodeURIComponent).join("/");

  const fileContent = fs.readFileSync(localFilePath);
  const uploadRes   = await fetch(
    `https://graph.microsoft.com/v1.0/sites/${siteId}/drive/root:/${encodedPath}:/content`,
    {
      method:  "PUT",
      headers: { ...headers, "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
      body:    fileContent,
    }
  );
  if (!uploadRes.ok) {
    const errText = await uploadRes.text();
    if (uploadRes.status === 423 || errText.includes("resourceLocked")) {
      throw new Error(`Arquivo bloqueado no ${label} — feche-o no Excel ou SharePoint Online e tente novamente.`);
    }
    throw new Error(`Erro no upload (${label}): ${errText}`);
  }

  const { webUrl } = await uploadRes.json();
  console.log(`  ✓ Atualizado no ${label}: ${webUrl}`);
}

// ── Limitador de concorrência ─────────────────────────────────────────────────
// Evita throttling da API Topdesk ao buscar histórico de muitos ativos em paralelo.
// Mantém no máximo `concurrency` requisições em andamento simultaneamente.
async function withConcurrency(items, concurrency, fn) {
  let index = 0;
  async function worker() {
    while (index < items.length) {
      const i = index++;
      await fn(items[i], i);
    }
  }
  const workers = Array.from({ length: Math.min(concurrency, items.length) }, worker);
  await Promise.all(workers);
}

// ─────────────────────────────────────────────────────────────────────────────

async function writeExcel(rows, outputPath) {
  const wb = new ExcelJS.Workbook();

  // Aba 1: Resumo executivo (abre primeiro)
  buildSummarySheet(wb, rows);

  // Aba 2: Detalhe completo
  const ws = wb.addWorksheet("Não Devolvidos");
  applySheetStyle(ws, rows);

  await wb.xlsx.writeFile(outputPath);
}

// ── Relatório de exauridos ───────────────────────────────────────────────────
// Lista dos ativos cujo CC não foi recuperado nem do campo do asset nem do
// histórico (esgotaram ou foram cortados pelo MAX_PAGES). Para que o time de
// administração possa preencher manualmente no Topdesk.

const EXAURIDOS_COLUMNS = [
  { header: "TIPO",                          key: "tipo" },
  { header: "NOME DO DISPOSITIVO",           key: "nome_dispositivo" },
  { header: "MODELO",                        key: "modelo" },
  { header: "PATRIMÔNIO",                    key: "patrimonio" },
  { header: "USUÁRIO ANTIGO",                key: "usuario" },
  { header: "UF",                            key: "uf" },
  { header: "DEPARTAMENTO",                  key: "departamento" },
  { header: "DIAS SEM RETORNO",              key: "dias_sem_retorno" },
  { header: "DATA MUDANÇA P/ NÃO DEVOLVIDO", key: "data_mudanca" },
  { header: "N° DO CHAMADO DA AUTOMAÇÃO",    key: "num_chamado" },
  { header: "MOTIVO",                        key: "motivo" },
];

function buildExauridosSheet(wb, rows) {
  const ws = wb.addWorksheet("Sem CC");
  const nCols = EXAURIDOS_COLUMNS.length;
  const lastCol = String.fromCharCode("A".charCodeAt(0) + nCols - 1); // "A" + 10 = "K"

  // Linha 1: título mesclado
  ws.mergeCells(`A1:${lastCol}1`);
  const titleCell = ws.getCell("A1");
  titleCell.value     = "ATIVOS SEM CENTRO DE CUSTO — PREENCHER NO TOPDESK";
  titleCell.font      = { bold: true, size: 14, color: { argb: PAL_WHITE }, name: "Calibri" };
  titleCell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: PAL_DARK } };
  titleCell.alignment = { horizontal: "center", vertical: "middle" };
  titleCell.border    = { bottom: { style: "thick", color: { argb: PAL_ORANGE } } };
  ws.getRow(1).height = 36;

  // Linha 2: cabeçalho das colunas
  const headerRow = ws.getRow(2);
  EXAURIDOS_COLUMNS.forEach((c, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value     = c.header;
    cell.font      = { bold: true, color: { argb: PAL_WHITE }, name: "Calibri", size: 10 };
    cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: PAL_DARK } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border    = { bottom: { style: "medium", color: { argb: PAL_ORANGE } } };
  });
  headerRow.height = 32;

  // Linhas 3+: dados
  const diasColIdx   = EXAURIDOS_COLUMNS.findIndex((c) => c.key === "dias_sem_retorno") + 1;
  const motivoColIdx = EXAURIDOS_COLUMNS.findIndex((c) => c.key === "motivo") + 1;

  rows.forEach((row, i) => {
    const r = ws.getRow(i + 3);
    EXAURIDOS_COLUMNS.forEach((c, ci) => { r.getCell(ci + 1).value = row[c.key]; });

    const rowBg = i % 2 === 1 ? PAL_ALT : PAL_WHITE;
    r.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      cell.font      = { color: { argb: PAL_TEXT }, name: "Calibri", size: 10 };
      cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: rowBg } };
      cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      cell.border    = { bottom: { style: "thin", color: { argb: "FFE0E0E0" } } };

      if (colNumber === diasColIdx) {
        const bucket = agingBucket(cell.value);
        if (bucket) {
          cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: bucket.bg } };
          cell.font      = { bold: true, color: { argb: bucket.text }, name: "Calibri", size: 10 };
          cell.alignment = { horizontal: "center", vertical: "middle" };
        }
      }

      if (colNumber === motivoColIdx) {
        cell.font      = { italic: true, color: { argb: "FF9C0006" }, name: "Calibri", size: 9 };
        cell.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
      }
    });
    r.height = 22;
  });

  // ── Nota de escopo (rodapé) ──────────────────────────────────────────────
  const notaRow = rows.length + 4;
  ws.mergeCells(`A${notaRow}:${lastCol}${notaRow}`);
  const nota = ws.getCell(`A${notaRow}`);
  nota.value =
    "Estes ativos aparecem no consolidado (Nao Devolvidos.xlsx), mas NÃO nos arquivos por obra — sem centro de custo não há " +
    "obra para alocá-los. Preencher o CC no TOPdesk os inclui automaticamente nos relatórios por obra a partir do dia seguinte.";
  nota.font      = { italic: true, size: 8, color: { argb: "FF999999" }, name: "Calibri" };
  nota.alignment = { horizontal: "left", vertical: "middle", wrapText: true };
  ws.getRow(notaRow).height = 30;

  // Largura por coluna (manual, baseada no header + maior valor visto)
  EXAURIDOS_COLUMNS.forEach((c, ci) => {
    let maxLen = c.header.length;
    rows.forEach((r) => {
      const val = r[c.key] != null ? String(r[c.key]) : "";
      if (val.length > maxLen) maxLen = val.length;
    });
    ws.getColumn(ci + 1).width = Math.min(maxLen + 4, 60);
  });

  // Auto-filtro na linha do cabeçalho + congela título + cabeçalho
  ws.autoFilter = { from: { row: 2, column: 1 }, to: { row: 2, column: nCols } };
  ws.views = [{ state: "frozen", ySplit: 2 }];
}

async function writeExauridosExcel(rows, outputPath) {
  const wb = new ExcelJS.Workbook();
  buildExauridosSheet(wb, rows);
  await wb.xlsx.writeFile(outputPath);
}

async function main() {
  const missing = Object.entries(REQUIRED_ENV).filter(([, v]) => !v).map(([k]) => k);
  if (missing.length) {
    throw new Error(`Variáveis de ambiente ausentes: ${missing.join(", ")} (ver .env.example)`);
  }

  console.time("TOTAL");

  console.log("Buscando ativos não devolvidos...");
  console.time("1. fetchAssets (notebook + celular)");
  const notebooks = await fetchAssets(TEMPLATE_NOTEBOOK, "Notebook");
  const celulares = await fetchAssets(TEMPLATE_CELULAR, "Celular");
  const allAssets = [...notebooks, ...celulares];
  console.timeEnd("1. fetchAssets (notebook + celular)");

  console.log(`  ${notebooks.length} notebook(s) | ${celulares.length} celular(es) encontrado(s)`);

  const deviceNames = allAssets
    .map((a) => a["nome-da-maquina"] || a["nome-do-dispositivo"] || "")
    .filter(Boolean);

  console.log("Buscando chamados de automação...");
  console.time("2. buildIncidentMap");
  const incidentMap = await buildIncidentMap(deviceNames);
  console.timeEnd("2. buildIncidentMap");
  console.log(`  ${Object.keys(incidentMap).length} chamado(s) vinculado(s)`);

  console.log("Buscando datas de mudança de status e centros de custo no histórico dos ativos...");
  console.time("3. getHistoryData (todos os ativos)");
  const historyDates = {};
  const historyCC    = {};
  // Limita a 8 requisições simultâneas ao /history/pastItems para evitar throttling da API.
  // Curto-circuito: se o asset já tem centro-de-custo preenchido, passa direto pro getHistoryData
  // (ele só vai procurar o status, sem varrer atrás de CC).
  await withConcurrency(allAssets, 8, async (asset) => {
    const assetId   = asset.unid || asset.id;
    const assetName = asset["nome-da-maquina"] || asset["nome-do-dispositivo"] || "";
    const knownCC   = asset["centro-de-custo"] ? String(asset["centro-de-custo"]).trim() : null;
    const { statusChangeDate, lastCentroCusto } = await getHistoryData(assetId, assetName, knownCC);
    if (statusChangeDate) historyDates[assetId] = statusChangeDate;
    if (lastCentroCusto)  historyCC[assetId]    = lastCentroCusto;
  });
  console.timeEnd("3. getHistoryData (todos os ativos)");
  console.log(`  ${Object.keys(historyDates).length} data(s) de status encontrada(s) no histórico`);
  console.log(`  ${Object.keys(historyCC).length} centro(s) de custo recuperado(s) do histórico`);

  // === Diagnóstico: top 15 ativos mais lentos no histórico ===
  console.log("\n  📊 Top 15 ativos mais lentos no /history/pastItems:");
  console.log("  " + "─".repeat(95));
  console.log("  " + ["páginas", "itens", "ms", "status?", "cc?", "esgot?", "ativo"].map((s, i) =>
    [8, 6, 7, 8, 5, 7, 0][i] ? s.padEnd([8, 6, 7, 8, 5, 7, 30][i]) : s).join(""));
  console.log("  " + "─".repeat(95));
  const top = [..._historyStats].sort((a, b) => b.ms - a.ms).slice(0, 15);
  for (const s of top) {
    console.log("  " + [
      String(s.pages).padEnd(8),
      String(s.items_seen).padEnd(6),
      String(s.ms).padEnd(7),
      (s.found_status ? "✓" : "✗").padEnd(8),
      (s.found_cc ? "✓" : "✗").padEnd(5),
      (s.exhausted ? "SIM" : "não").padEnd(7),
      `${s.assetName} (${s.assetId.slice(0, 8)})`,
    ].join(""));
  }

  // Agregados
  const totalPages = _historyStats.reduce((sum, s) => sum + s.pages, 0);
  const totalItems = _historyStats.reduce((sum, s) => sum + s.items_seen, 0);
  const exhausted  = _historyStats.filter((s) => s.exhausted).length;
  const semStatus  = _historyStats.filter((s) => !s.found_status).length;
  const semCC      = _historyStats.filter((s) => !s.found_cc).length;
  console.log("  " + "─".repeat(95));
  console.log(`  Resumo: ${_historyStats.length} ativos | ${totalPages} páginas (~${(totalPages / _historyStats.length).toFixed(1)}/ativo) | ${totalItems} itens lidos`);
  console.log(`          ${exhausted} esgotaram histórico | ${semStatus} sem status encontrado | ${semCC} sem CC encontrado`);
  console.log();

  const DATE_FILTER_FROM = new Date("2026-01-01T00:00:00.000Z");
  const NOW              = Date.now();

  const rowsRaw = allAssets
    .map((asset) => {
      const nome    = asset["nome-da-maquina"] || asset["nome-do-dispositivo"] || "";
      const assetId = asset.unid || asset.id;
      const ufMatch = nome.match(/^[A-Z]{2}-([A-Z]{2})-/);
      const rawDate = historyDates[assetId] || incidentMap[nome]?.date || "";
      const dias    = rawDate ? Math.floor((NOW - new Date(rawDate)) / 86_400_000) : "";

      return {
        nome_dispositivo:   nome,
        modelo:             asset["modelo"] || "",
        tipo:               asset._tipo,
        patrimonio:         asset["numero-de-patrimonio"] || "",
        usuario:            asset["dsp"] || asset["usuario"] || "",
        centro_custo:       asset["centro-de-custo"] || historyCC[assetId] || "",
        uf:                 ufMatch ? ufMatch[1] : "",
        departamento:       asset["departamento"] != null ? String(asset["departamento"]) : "",
        dias_sem_retorno:   dias,
        data_mudanca:       rawDate ? formatDate(rawDate) : "",
        data_nao_devolvido: asset["data-nao-devolvido"] ? formatDate(asset["data-nao-devolvido"]) : "",
        num_chamado:        incidentMap[nome]?.number || "",
        _rawDate:           rawDate,
        _assetId:           assetId,
      };
    })
    .filter((row) => {
      if (!row._rawDate) return true;
      return new Date(row._rawDate) >= DATE_FILTER_FROM;
    })
    // Ordenar do mais antigo (mais crítico) para o mais recente
    .sort((a, b) => {
      const da = typeof a.dias_sem_retorno === "number" ? a.dias_sem_retorno : -1;
      const db = typeof b.dias_sem_retorno === "number" ? b.dias_sem_retorno : -1;
      return db - da;
    });

  // Enriquecer UF a partir do usuário antigo (filial da pessoa no Topdesk)
  // para os devices cujo nome não segue o padrão XX-UF-...
  const semUf = rowsRaw.filter((r) => !r.uf && r.usuario);
  if (semUf.length) {
    console.log(`Recuperando UF de ${semUf.length} device(s) pelo usuário antigo...`);
    console.time("4. getUserUf (lookup de pessoas)");
    await withConcurrency(semUf, 8, async (r) => {
      r.uf = await getUserUf(r.usuario);
    });
    console.timeEnd("4. getUserUf (lookup de pessoas)");
    const recuperadas = semUf.filter((r) => r.uf).length;
    console.log(`  ${recuperadas} UF(s) recuperada(s) via usuário`);
  }

  const rows = rowsRaw.map(({ _rawDate, _assetId, ...rest }) => rest);

  // Remove arquivos xlsx anteriores
  const dir = __dirname;
  for (const file of fs.readdirSync(dir)) {
    if (/^nao_devolvidos_.*\.xlsx$/.test(file)) {
      try {
        fs.unlinkSync(path.join(dir, file));
        console.log(`  Arquivo anterior removido: ${file}`);
      } catch (e) {
        console.warn(`  Não foi possível remover ${file}: ${e.message}`);
      }
    }
  }

  console.time("5. writeExcel (geração do .xlsx)");
  const today = new Date().toISOString().slice(0, 10);
  let output  = `nao_devolvidos_${today}.xlsx`;
  let attempt = 1;
  while (true) {
    try {
      await writeExcel(rows, output);
      break;
    } catch (err) {
      if (err.code === "EBUSY" && attempt <= 9) {
        output = `nao_devolvidos_${today}_${attempt}.xlsx`;
        attempt++;
      } else {
        throw err;
      }
    }
  }
  console.timeEnd("5. writeExcel (geração do .xlsx)");
  console.log(`\nArquivo gerado: ${output} (${rows.length} linha(s))`);

  console.time("6. uploadToSharePoint");
  await uploadToSharePoint(path.resolve(output), { label: "SharePoint principal" });
  console.timeEnd("6. uploadToSharePoint");

  // ── Relatório de exauridos ───────────────────────────────────────────────
  // Ativos cujo CC ficou vazio mesmo após buscar no histórico. O motivo de
  // cada um vem do _historyStats coletado em getHistoryData.
  const statsByAssetId = new Map(_historyStats.map((s) => [s.assetId, s]));
  const exauridosRows  = rowsRaw
    .filter((r) => !r.centro_custo)
    .map((r) => {
      const stats = statsByAssetId.get(r._assetId);
      let motivo;
      if (!stats)                       motivo = "Histórico não consultado";
      else if (stats.hit_max_pages)     motivo = `Cortado no limite de ${MAX_HISTORY_PAGES} páginas`;
      else if (stats.exhausted)         motivo = `Histórico esgotado sem evento de CC (${stats.pages} págs)`;
      else if (!stats.found_cc)         motivo = `Sem evento de CC no histórico (${stats.pages} págs lidas)`;
      else                              motivo = "(sem motivo identificado)";

      return {
        tipo:               r.tipo,
        nome_dispositivo:   r.nome_dispositivo,
        modelo:             r.modelo,
        patrimonio:         r.patrimonio,
        usuario:            r.usuario,
        uf:                 r.uf,
        departamento:       r.departamento,
        dias_sem_retorno:   r.dias_sem_retorno,
        data_mudanca:       r.data_mudanca,
        num_chamado:        r.num_chamado,
        motivo,
      };
    });

  if (exauridosRows.length === 0) {
    console.log("\n  [exauridos] Nenhum ativo sem CC — relatório não gerado.");
  } else {
    console.log(`\n  [exauridos] ${exauridosRows.length} ativo(s) sem CC — gerando relatório separado...`);
    console.time("7. writeExauridos + upload");

    const exOut = `nao_devolvidos_sem_cc_${today}.xlsx`;
    let exAttempt = 1;
    let exFinal   = exOut;
    while (true) {
      try {
        await writeExauridosExcel(exauridosRows, exFinal);
        break;
      } catch (err) {
        if (err.code === "EBUSY" && exAttempt <= 9) {
          exFinal = `nao_devolvidos_sem_cc_${today}_${exAttempt}.xlsx`;
          exAttempt++;
        } else {
          throw err;
        }
      }
    }
    console.log(`  ✓ Arquivo local gerado: ${exFinal}`);

    await uploadToSharePoint(path.resolve(exFinal), {
      siteUrl:  SHAREPOINT_EXAURIDOS_SITE_URL,
      folder:   SHAREPOINT_EXAURIDOS_FOLDER,
      filename: SHAREPOINT_EXAURIDOS_FILENAME,
      label:    "SharePoint (sem CC)",
    });
    console.timeEnd("7. writeExauridos + upload");
  }

  console.timeEnd("TOTAL");
}

main().catch((err) => {
  console.error("Erro:", err.message);
  process.exit(1);
});
