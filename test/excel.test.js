const test = require("node:test");
const assert = require("node:assert/strict");
const fs = require("fs");
const os = require("os");
const path = require("path");
const ExcelJS = require("exceljs");

const { agingBucket } = require("../src/excel/estilo");
const { isTablet, summarize } = require("../src/excel/resumo");
const { writeExcel, writeSemCCExcel } = require("../src/excel");
const { COLUMNS } = require("../src/excel/detalhe");

const row = (over) => ({ tipo: "Notebook", nome_dispositivo: "NB-SP-0001", uf: "SP", centro_custo: "1101", dias_sem_retorno: 10, ...over });

test("agingBucket separa em ≤30, 31–90 e >90 dias", () => {
  assert.match(agingBucket(0).label, /Em dia/);
  assert.match(agingBucket(30).label, /Em dia/);
  assert.match(agingBucket(31).label, /Atenção/);
  assert.match(agingBucket(90).label, /Atenção/);
  assert.match(agingBucket(91).label, /Crítico/);
  assert.equal(agingBucket(""), null);
});

test("isTablet reconhece tablet pelo nome, só no template de celular", () => {
  assert.equal(isTablet(row({ tipo: "Celular", nome_dispositivo: "CL-SP-Tablet-3" })), true);
  assert.equal(isTablet(row({ tipo: "Celular", nome_dispositivo: "CL-SP-0003" })), false);
  assert.equal(isTablet(row({ tipo: "Notebook", nome_dispositivo: "NB-TABLET" })), false);
});

test("summarize conta KPIs, faixas, UFs e centros de custo", () => {
  const s = summarize([
    row({ dias_sem_retorno: 5 }),
    row({ dias_sem_retorno: 45, uf: "RJ" }),
    row({ tipo: "Celular", nome_dispositivo: "CL-RJ-0002", dias_sem_retorno: 120, uf: "RJ", centro_custo: "" }),
    row({ tipo: "Celular", nome_dispositivo: "CL-MG-TABLET-1", dias_sem_retorno: "", uf: "" }),
  ]);

  assert.deepEqual([s.total, s.notebooks, s.celulares, s.tablets], [4, 2, 1, 1]);
  assert.deepEqual(s.agingStats.map((a) => a.count), [1, 1, 1]);
  assert.equal(s.semData, 1);
  assert.deepEqual(s.ufEntries, [["RJ", 2], ["SP", 1], ["Outros", 1]]);
  assert.deepEqual(s.ccEntries, [["1101", 2, 0, 1, 3], ["(Não informado)", 0, 1, 0, 1]]);
});

test("writeExcel gera as abas Resumo e Não Devolvidos", async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "nao-devolvidos-"));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const file = path.join(dir, "relatorio.xlsx");

  await writeExcel([row({ dias_sem_retorno: 100 }), row({ nome_dispositivo: "NB-RJ-0002", uf: "RJ" })], file);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  assert.deepEqual(wb.worksheets.map((ws) => ws.name), ["Resumo", "Não Devolvidos"]);

  const resumo = wb.getWorksheet("Resumo");
  assert.equal(resumo.getCell("A1").value, "NÃO DEVOLVIDOS — RESUMO EXECUTIVO");
  assert.equal(resumo.getCell("A5").value, 2);

  const detalhe = wb.getWorksheet("Não Devolvidos");
  assert.equal(detalhe.getRow(1).getCell(1).value, COLUMNS[0].header);
  assert.equal(detalhe.rowCount, 3);
  const dias = detalhe.getRow(2).getCell(COLUMNS.findIndex((c) => c.key === "dias_sem_retorno") + 1);
  assert.equal(dias.value, 100);
  assert.equal(dias.fill.fgColor.argb, agingBucket(100).bg);
});

test("writeSemCCExcel gera a aba Sem CC com o motivo", async (t) => {
  const dir = fs.mkdtempSync(path.join(os.tmpdir(), "nao-devolvidos-"));
  t.after(() => fs.rmSync(dir, { recursive: true, force: true }));
  const file = path.join(dir, "sem-cc.xlsx");

  await writeSemCCExcel([{ tipo: "Notebook", nome_dispositivo: "NB-SP-0001", motivo: "Cortado no limite de 100 páginas" }], file);

  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(file);
  const ws = wb.getWorksheet("Sem CC");
  assert.equal(ws.getCell("A1").value, "ATIVOS SEM CENTRO DE CUSTO — PREENCHER NO TOPDESK");
  assert.equal(ws.getCell("K3").value, "Cortado no limite de 100 páginas");
});
