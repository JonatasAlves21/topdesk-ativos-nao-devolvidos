const test = require("node:test");
const assert = require("node:assert/strict");

const {
  deviceName, ufFromDeviceName, diasDesde, applyHistoryItems, matchIncidents,
  buildRows, stripInternal, motivoSemCC, buildSemCCRows,
} = require("../src/regras");

const DAY = 86_400_000;
const NOW = Date.parse("2026-06-30T12:00:00Z");

test("deviceName usa o campo de notebook ou o de celular", () => {
  assert.equal(deviceName({ "nome-da-maquina": "NB-SP-0001" }), "NB-SP-0001");
  assert.equal(deviceName({ "nome-do-dispositivo": "CL-RJ-0002" }), "CL-RJ-0002");
  assert.equal(deviceName({}), "");
});

test("ufFromDeviceName extrai a UF do padrão XX-UF-...", () => {
  assert.equal(ufFromDeviceName("NB-SP-0412"), "SP");
  assert.equal(ufFromDeviceName("CL-BA-TABLET-3"), "BA");
  assert.equal(ufFromDeviceName("NOTE-ANTIGO-9"), "");
  assert.equal(ufFromDeviceName("nb-sp-0412"), "");
  assert.equal(ufFromDeviceName(undefined), "");
});

test("diasDesde conta dias inteiros e devolve vazio sem data", () => {
  assert.equal(diasDesde(new Date(NOW - 45 * DAY).toISOString(), NOW), 45);
  assert.equal(diasDesde(new Date(NOW - 45 * DAY + 1000).toISOString(), NOW), 44);
  assert.equal(diasDesde("", NOW), "");
});

test("applyHistoryItems acha a data do status e o CC pelo oldValue", () => {
  const found = { statusChangeDate: null, lastCentroCusto: null };
  const done = applyHistoryItems([
    { date: "2026-03-01", field: { fieldName: "modelo", newValue: "x" } },
    { date: "2026-03-02", field: { fieldName: "status", newValue: "OUTRO" } },
    { date: "2026-03-03", field: { fieldName: "status", newValue: "S_ND" } },
    { date: "2026-03-03", field: { fieldName: "centro-de-custo", newValue: "", oldValue: "1101 - Obra Alfa" } },
  ], found, "S_ND");

  assert.equal(done, true);
  assert.deepEqual(found, { statusChangeDate: "2026-03-03", lastCentroCusto: "1101 - Obra Alfa" });
});

test("applyHistoryItems não sobrescreve o que já foi encontrado e pede a próxima página", () => {
  const found = { statusChangeDate: null, lastCentroCusto: "CC do ativo" };
  const done = applyHistoryItems([
    { date: "2026-01-10", field: { fieldName: "centro-de-custo", newValue: "outro CC" } },
  ], found, "S_ND");

  assert.equal(done, false);
  assert.equal(found.lastCentroCusto, "CC do ativo");
  assert.equal(found.statusChangeDate, null);
});

test("matchIncidents vincula o chamado mais antigo que cita o dispositivo", () => {
  const map = matchIncidents([
    { number: "I-3", creationDate: "2026-04-01", request: "NB-SP-0001 de novo" },
    { number: "I-1", creationDate: "2026-02-01", request: "Ativo NB-SP-0001 não devolvido" },
    { number: "I-2", creationDate: "2026-03-01", request: "Ativo CL-RJ-0002 não devolvido" },
  ], ["NB-SP-0001", "CL-RJ-0002", "NB-MG-0003"]);

  assert.deepEqual(map, {
    "NB-SP-0001": { number: "I-1", date: "2026-02-01" },
    "CL-RJ-0002": { number: "I-2", date: "2026-03-01" },
  });
});

test("buildRows monta, filtra pela data de corte e ordena do mais antigo", () => {
  const assets = [
    { unid: "a1", _tipo: "Notebook", "nome-da-maquina": "NB-SP-0001", dsp: "Ana", "centro-de-custo": "" },
    { unid: "a2", _tipo: "Celular", "nome-do-dispositivo": "CL-RJ-0002", usuario: "Bia", "centro-de-custo": "2002", departamento: 7 },
    { unid: "a3", _tipo: "Notebook", "nome-da-maquina": "NB-MG-0003" },
    { unid: "a4", _tipo: "Notebook", "nome-da-maquina": "NB-PR-0004" },
  ];
  const rows = buildRows(assets, {
    historyDates: { a1: new Date(NOW - 10 * DAY).toISOString(), a4: "2025-12-01T00:00:00Z" },
    historyCC:    { a1: "1001 - Hist" },
    incidentMap:  { "CL-RJ-0002": { number: "I-9", date: new Date(NOW - 100 * DAY).toISOString() } },
    now: NOW,
  });

  // a4 sai (mudou antes de 2026); a3 fica (sem data) e vai para o fim
  assert.deepEqual(rows.map((r) => r._assetId), ["a2", "a1", "a3"]);

  const [cel, nb, semData] = rows;
  assert.equal(cel.dias_sem_retorno, 100);           // data veio do chamado
  assert.equal(cel.num_chamado, "I-9");
  assert.equal(cel.centro_custo, "2002");             // CC do próprio ativo
  assert.equal(cel.departamento, "7");
  assert.equal(cel.uf, "RJ");
  assert.equal(nb.centro_custo, "1001 - Hist");       // CC recuperado do histórico
  assert.equal(nb.usuario, "Ana");
  assert.equal(semData.dias_sem_retorno, "");
  assert.equal(semData.data_mudanca, "");
});

test("stripInternal remove só os campos internos", () => {
  assert.deepEqual(stripInternal([{ a: 1, _rawDate: "x", _assetId: "y" }]), [{ a: 1 }]);
});

test("motivoSemCC explica por que o CC não foi recuperado", () => {
  assert.equal(motivoSemCC(undefined, 100), "Histórico não consultado");
  assert.equal(motivoSemCC({ hit_max_pages: true, pages: 100 }, 100), "Cortado no limite de 100 páginas");
  assert.equal(motivoSemCC({ exhausted: true, pages: 4 }, 100), "Histórico esgotado sem evento de CC (4 págs)");
  assert.equal(motivoSemCC({ found_cc: false, pages: 7 }, 100), "Sem evento de CC no histórico (7 págs lidas)");
  assert.equal(motivoSemCC({ found_cc: true, pages: 2 }, 100), "(sem motivo identificado)");
});

test("buildSemCCRows lista só quem ficou sem CC, com o motivo", () => {
  const rows = [
    { _assetId: "a1", tipo: "Notebook", nome_dispositivo: "NB-SP-0001", centro_custo: "" },
    { _assetId: "a2", tipo: "Celular", nome_dispositivo: "CL-RJ-0002", centro_custo: "2002" },
  ];
  const semCC = buildSemCCRows(rows, [{ assetId: "a1", hit_max_pages: true, pages: 100 }], 100);

  assert.equal(semCC.length, 1);
  assert.equal(semCC[0].nome_dispositivo, "NB-SP-0001");
  assert.equal(semCC[0].motivo, "Cortado no limite de 100 páginas");
  assert.equal("_assetId" in semCC[0], false);
});
