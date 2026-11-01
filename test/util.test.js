const test = require("node:test");
const assert = require("node:assert/strict");

const { fetchWithRetry, withConcurrency, formatDate, writeAvoidingLock } = require("../src/util");
const { missingEnv } = require("../src/config");

test("formatDate usa dd/mm/aaaa hh:mm e aceita vazio", () => {
  assert.equal(formatDate(new Date(2026, 0, 5, 9, 7)), "05/01/2026 09:07");
  assert.equal(formatDate(""), "");
  assert.equal(formatDate(null), "");
});

test("withConcurrency processa todos os itens sem passar do limite", async () => {
  let running = 0;
  let peak = 0;
  const seen = [];
  await withConcurrency([...Array(20).keys()], 3, async (n) => {
    running++;
    peak = Math.max(peak, running);
    await new Promise((r) => setTimeout(r, 2));
    seen.push(n);
    running--;
  });
  assert.equal(peak, 3);
  assert.deepEqual(seen.sort((a, b) => a - b), [...Array(20).keys()]);
});

test("withConcurrency com lista vazia termina sem chamar a função", async () => {
  let calls = 0;
  await withConcurrency([], 8, async () => { calls++; });
  assert.equal(calls, 0);
});

test("fetchWithRetry repete falhas de rede e devolve a resposta", async (t) => {
  let calls = 0;
  t.mock.method(globalThis, "fetch", async () => {
    calls++;
    if (calls < 3) throw new TypeError("fetch failed");
    return { ok: true };
  });
  const res = await fetchWithRetry("https://exemplo.test", {}, 4, 1);
  assert.equal(res.ok, true);
  assert.equal(calls, 3);
});

test("fetchWithRetry desiste depois das tentativas", async (t) => {
  t.mock.method(globalThis, "fetch", async () => { throw new TypeError("fetch failed"); });
  await assert.rejects(fetchWithRetry("https://exemplo.test", {}, 2, 1), /fetch failed/);
  assert.equal(globalThis.fetch.mock.callCount(), 3);
});

test("writeAvoidingLock troca o nome quando o arquivo está aberto (EBUSY)", async () => {
  const tried = [];
  const name = await writeAvoidingLock("relatorio", async (file) => {
    tried.push(file);
    if (tried.length < 3) throw Object.assign(new Error("busy"), { code: "EBUSY" });
  });
  assert.equal(name, "relatorio_2.xlsx");
  assert.deepEqual(tried, ["relatorio.xlsx", "relatorio_1.xlsx", "relatorio_2.xlsx"]);
});

test("writeAvoidingLock não engole outros erros", async () => {
  await assert.rejects(
    writeAvoidingLock("relatorio", async () => { throw Object.assign(new Error("negado"), { code: "EACCES" }); }),
    /negado/,
  );
});

test("missingEnv lista as variáveis vazias", () => {
  assert.deepEqual(missingEnv({ A: "1", B: "", C: undefined }), ["B", "C"]);
});
