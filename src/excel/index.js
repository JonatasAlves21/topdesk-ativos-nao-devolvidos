// Gravação dos dois arquivos Excel do relatório.

const ExcelJS = require("exceljs");
const { buildSummarySheet } = require("./resumo");
const { buildDetailSheet } = require("./detalhe");
const { buildSemCCSheet } = require("./sem-cc");

// Relatório principal: aba Resumo (abre primeiro) + aba de detalhe.
async function writeExcel(rows, outputPath) {
  const wb = new ExcelJS.Workbook();
  buildSummarySheet(wb, rows);
  buildDetailSheet(wb, rows);
  await wb.xlsx.writeFile(outputPath);
}

// Relatório dos ativos sem centro de custo.
async function writeSemCCExcel(rows, outputPath) {
  const wb = new ExcelJS.Workbook();
  buildSemCCSheet(wb, rows);
  await wb.xlsx.writeFile(outputPath);
}

module.exports = { writeExcel, writeSemCCExcel };
