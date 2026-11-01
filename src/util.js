// Utilitários sem regra de negócio: retry de rede, concorrência, datas e gravação.

// fetch com retry + backoff — tolera falhas de rede transitórias ("fetch failed")
async function fetchWithRetry(url, options, retries = 4, baseDelayMs = 500) {
  for (let attempt = 0; attempt <= retries; attempt++) {
    try {
      return await fetch(url, options);
    } catch (err) {
      if (attempt === retries) throw err;
      const wait = baseDelayMs * Math.pow(2, attempt); // 500ms, 1s, 2s, 4s
      await new Promise((r) => setTimeout(r, wait));
    }
  }
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

// dd/mm/aaaa hh:mm no fuso local
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

// Grava `<baseName>.xlsx`. Se o arquivo estiver aberto no Excel (EBUSY), tenta
// `<baseName>_1.xlsx` … `<baseName>_9.xlsx`. Devolve o nome usado.
async function writeAvoidingLock(baseName, write) {
  let output = `${baseName}.xlsx`;
  for (let attempt = 1; ; attempt++) {
    try {
      await write(output);
      return output;
    } catch (err) {
      if (err.code !== "EBUSY" || attempt > 9) throw err;
      output = `${baseName}_${attempt}.xlsx`;
    }
  }
}

module.exports = { fetchWithRetry, withConcurrency, formatDate, writeAvoidingLock };
