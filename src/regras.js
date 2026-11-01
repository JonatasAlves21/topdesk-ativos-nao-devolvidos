// Regras de negócio puras: nada aqui chama a API nem grava arquivo, por isso
// tudo é coberto pelos testes em test/regras.test.js.

const { formatDate } = require("./util");

// O relatório considera só mudanças de status a partir desta data.
const DATE_FILTER_FROM = new Date("2026-01-01T00:00:00.000Z");

const DAY_MS = 86_400_000;

// Notebooks usam "nome-da-maquina"; celulares e tablets, "nome-do-dispositivo".
function deviceName(asset) {
  return asset["nome-da-maquina"] || asset["nome-do-dispositivo"] || "";
}

// Padrão de nome XX-UF-... (ex.: NB-SP-0412 → "SP"). Fora do padrão, devolve "".
function ufFromDeviceName(nome) {
  const ufMatch = (nome || "").match(/^[A-Z]{2}-([A-Z]{2})-/);
  return ufMatch ? ufMatch[1] : "";
}

function diasDesde(rawDate, now) {
  return rawDate ? Math.floor((now - new Date(rawDate)) / DAY_MS) : "";
}

// Processa uma página do /history/pastItems, preenchendo `found` com:
//   - statusChangeDate: data em que o status mudou para "Não devolvido"
//   - lastCentroCusto:  último centro de custo registrado
// Devolve true quando os dois já foram encontrados e a varredura pode parar.
function applyHistoryItems(items, found, statusNaoDev) {
  for (const item of items) {
    const fn = item.field?.fieldName;

    if (!found.statusChangeDate && fn === "status" && item.field?.newValue === statusNaoDev) {
      found.statusChangeDate = item.date;
    }

    if (!found.lastCentroCusto && fn === "centro-de-custo") {
      // A automação zera o campo ao marcar o ativo como "Não devolvido";
      // o oldValue preserva o último valor real.
      const val = item.field?.newValue || item.field?.oldValue;
      if (val) found.lastCentroCusto = String(val);
    }

    if (found.statusChangeDate && found.lastCentroCusto) break;
  }
  return !!(found.statusChangeDate && found.lastCentroCusto);
}

// Associa cada dispositivo ao chamado MAIS ANTIGO da automação que cita o nome dele.
function matchIncidents(incidents, deviceNames) {
  const byDate = [...incidents].sort((a, b) => new Date(a.creationDate) - new Date(b.creationDate));

  const incidentMap = {};
  for (const incident of byDate) {
    const body = incident.request || "";
    for (const name of deviceNames) {
      if (!incidentMap[name] && body.includes(name)) {
        incidentMap[name] = { number: incident.number, date: incident.creationDate };
      }
    }
  }
  return incidentMap;
}

// Monta as linhas do relatório a partir dos ativos e do que veio do histórico e
// dos chamados. Aplica o filtro de data e ordena do mais antigo (mais crítico)
// para o mais recente. Os campos com "_" são internos e saem antes do Excel.
function buildRows(assets, { historyDates = {}, historyCC = {}, incidentMap = {}, now = Date.now(), dateFrom = DATE_FILTER_FROM } = {}) {
  return assets
    .map((asset) => {
      const nome    = deviceName(asset);
      const assetId = asset.unid || asset.id;
      const rawDate = historyDates[assetId] || incidentMap[nome]?.date || "";

      return {
        nome_dispositivo:   nome,
        modelo:             asset["modelo"] || "",
        tipo:               asset._tipo,
        patrimonio:         asset["numero-de-patrimonio"] || "",
        usuario:            asset["dsp"] || asset["usuario"] || "",
        centro_custo:       asset["centro-de-custo"] || historyCC[assetId] || "",
        uf:                 ufFromDeviceName(nome),
        departamento:       asset["departamento"] != null ? String(asset["departamento"]) : "",
        dias_sem_retorno:   diasDesde(rawDate, now),
        data_mudanca:       rawDate ? formatDate(rawDate) : "",
        data_nao_devolvido: asset["data-nao-devolvido"] ? formatDate(asset["data-nao-devolvido"]) : "",
        num_chamado:        incidentMap[nome]?.number || "",
        _rawDate:           rawDate,
        _assetId:           assetId,
      };
    })
    .filter((row) => {
      if (!row._rawDate) return true;
      return new Date(row._rawDate) >= dateFrom;
    })
    .sort((a, b) => {
      const da = typeof a.dias_sem_retorno === "number" ? a.dias_sem_retorno : -1;
      const db = typeof b.dias_sem_retorno === "number" ? b.dias_sem_retorno : -1;
      return db - da;
    });
}

function stripInternal(rows) {
  return rows.map(({ _rawDate, _assetId, ...rest }) => rest);
}

// Por que o centro de custo não foi recuperado, a partir das estatísticas do histórico.
function motivoSemCC(stats, maxPages) {
  if (!stats)               return "Histórico não consultado";
  if (stats.hit_max_pages)  return `Cortado no limite de ${maxPages} páginas`;
  if (stats.exhausted)      return `Histórico esgotado sem evento de CC (${stats.pages} págs)`;
  if (!stats.found_cc)      return `Sem evento de CC no histórico (${stats.pages} págs lidas)`;
  return "(sem motivo identificado)";
}

// Linhas do relatório "sem CC": ativos cujo CC ficou vazio mesmo após o histórico.
function buildSemCCRows(rows, historyStats, maxPages) {
  const statsByAssetId = new Map(historyStats.map((s) => [s.assetId, s]));
  return rows
    .filter((r) => !r.centro_custo)
    .map((r) => ({
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
      motivo:             motivoSemCC(statsByAssetId.get(r._assetId), maxPages),
    }));
}

module.exports = {
  DATE_FILTER_FROM,
  deviceName,
  ufFromDeviceName,
  diasDesde,
  applyHistoryItems,
  matchIncidents,
  buildRows,
  stripInternal,
  motivoSemCC,
  buildSemCCRows,
};
