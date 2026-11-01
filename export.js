// Ponto de entrada: busca os ativos "Não devolvido" no TOPdesk, gera os relatórios
// em Excel e publica no SharePoint. Cada etapa é cronometrada no log.
//
//   1. ativos       src/topdesk.js   fetchAssets
//   2. chamados     src/topdesk.js   buildIncidentMap
//   3. histórico    src/topdesk.js   getHistoryData (8 em paralelo, teto de páginas)
//   4. UF           src/topdesk.js   getUserUf (só quando o nome não traz a UF)
//   5. Excel        src/excel/       writeExcel
//   6. SharePoint   src/sharepoint.js
//   7. sem CC       src/excel/       writeSemCCExcel + upload

const fs = require("fs");
const path = require("path");

const config = require("./src/config");
const {
  MAX_HISTORY_PAGES, historyStats, fetchAssets, getHistoryData, getUserUf, buildIncidentMap,
} = require("./src/topdesk");
const { deviceName, buildRows, stripInternal, buildSemCCRows } = require("./src/regras");
const { withConcurrency, writeAvoidingLock } = require("./src/util");
const { writeExcel, writeSemCCExcel } = require("./src/excel");
const { uploadToSharePoint } = require("./src/sharepoint");
const { printHistoryReport } = require("./src/diagnostico");

async function main() {
  const missing = config.missingEnv();
  if (missing.length) {
    throw new Error(`Variáveis de ambiente ausentes: ${missing.join(", ")} (ver .env.example)`);
  }

  console.time("TOTAL");

  console.log("Buscando ativos não devolvidos...");
  console.time("1. fetchAssets (notebook + celular)");
  const notebooks = await fetchAssets(config.TEMPLATE_NOTEBOOK, "Notebook");
  const celulares = await fetchAssets(config.TEMPLATE_CELULAR, "Celular");
  const allAssets = [...notebooks, ...celulares];
  console.timeEnd("1. fetchAssets (notebook + celular)");

  console.log(`  ${notebooks.length} notebook(s) | ${celulares.length} celular(es) encontrado(s)`);

  const deviceNames = allAssets.map(deviceName).filter(Boolean);

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
    const knownCC   = asset["centro-de-custo"] ? String(asset["centro-de-custo"]).trim() : null;
    const { statusChangeDate, lastCentroCusto } = await getHistoryData(assetId, deviceName(asset), knownCC);
    if (statusChangeDate) historyDates[assetId] = statusChangeDate;
    if (lastCentroCusto)  historyCC[assetId]    = lastCentroCusto;
  });
  console.timeEnd("3. getHistoryData (todos os ativos)");
  console.log(`  ${Object.keys(historyDates).length} data(s) de status encontrada(s) no histórico`);
  console.log(`  ${Object.keys(historyCC).length} centro(s) de custo recuperado(s) do histórico`);

  printHistoryReport(historyStats);

  const rowsRaw = buildRows(allAssets, { historyDates, historyCC, incidentMap, now: Date.now() });

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

  const rows = stripInternal(rowsRaw);

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
  const today  = new Date().toISOString().slice(0, 10);
  const output = await writeAvoidingLock(`nao_devolvidos_${today}`, (file) => writeExcel(rows, file));
  console.timeEnd("5. writeExcel (geração do .xlsx)");
  console.log(`\nArquivo gerado: ${output} (${rows.length} linha(s))`);

  console.time("6. uploadToSharePoint");
  await uploadToSharePoint(path.resolve(output), { label: "SharePoint principal" });
  console.timeEnd("6. uploadToSharePoint");

  // ── Relatório sem centro de custo ────────────────────────────────────────
  // Ativos cujo CC ficou vazio mesmo após buscar no histórico. O motivo de
  // cada um vem do historyStats coletado em getHistoryData.
  const semCCRows = buildSemCCRows(rowsRaw, historyStats, MAX_HISTORY_PAGES);

  if (semCCRows.length === 0) {
    console.log("\n  [exauridos] Nenhum ativo sem CC — relatório não gerado.");
  } else {
    console.log(`\n  [exauridos] ${semCCRows.length} ativo(s) sem CC — gerando relatório separado...`);
    console.time("7. writeExauridos + upload");

    const semCCFile = await writeAvoidingLock(`nao_devolvidos_sem_cc_${today}`, (file) => writeSemCCExcel(semCCRows, file));
    console.log(`  ✓ Arquivo local gerado: ${semCCFile}`);

    await uploadToSharePoint(path.resolve(semCCFile), {
      siteUrl:  config.SHAREPOINT_EXAURIDOS_SITE_URL,
      folder:   config.SHAREPOINT_EXAURIDOS_FOLDER,
      filename: config.SHAREPOINT_EXAURIDOS_FILENAME,
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
