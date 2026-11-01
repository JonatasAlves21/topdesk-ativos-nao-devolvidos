// ── Relatório sem centro de custo ────────────────────────────────────────────
// Lista dos ativos cujo CC não foi recuperado nem do campo do asset nem do
// histórico (esgotaram ou foram cortados pelo MAX_HISTORY_PAGES). Para que o
// time de administração possa preencher manualmente no Topdesk.

const { PAL_DARK, PAL_ORANGE, PAL_ALT, PAL_WHITE, PAL_TEXT, agingBucket } = require("./estilo");

const SEM_CC_COLUMNS = [
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

function buildSemCCSheet(wb, rows) {
  const ws = wb.addWorksheet("Sem CC");
  const nCols = SEM_CC_COLUMNS.length;
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
  SEM_CC_COLUMNS.forEach((c, i) => {
    const cell = headerRow.getCell(i + 1);
    cell.value     = c.header;
    cell.font      = { bold: true, color: { argb: PAL_WHITE }, name: "Calibri", size: 10 };
    cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: PAL_DARK } };
    cell.alignment = { horizontal: "center", vertical: "middle", wrapText: true };
    cell.border    = { bottom: { style: "medium", color: { argb: PAL_ORANGE } } };
  });
  headerRow.height = 32;

  // Linhas 3+: dados
  const diasColIdx   = SEM_CC_COLUMNS.findIndex((c) => c.key === "dias_sem_retorno") + 1;
  const motivoColIdx = SEM_CC_COLUMNS.findIndex((c) => c.key === "motivo") + 1;

  rows.forEach((row, i) => {
    const r = ws.getRow(i + 3);
    SEM_CC_COLUMNS.forEach((c, ci) => { r.getCell(ci + 1).value = row[c.key]; });

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
  SEM_CC_COLUMNS.forEach((c, ci) => {
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

module.exports = { SEM_CC_COLUMNS, buildSemCCSheet };
