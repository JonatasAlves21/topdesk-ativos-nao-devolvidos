// Aba "Não Devolvidos": uma linha por ativo, com filtro e cabeçalho congelado.

const { PAL_DARK, PAL_ORANGE, PAL_ALT, PAL_WHITE, PAL_TEXT, agingBucket } = require("./estilo");

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

function buildDetailSheet(wb, rows) {
  const ws = wb.addWorksheet("Não Devolvidos");
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

module.exports = { COLUMNS, buildDetailSheet };
