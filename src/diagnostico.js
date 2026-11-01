// Diagnóstico da varredura do histórico: os ativos mais lentos e os totais.
// Foi com estes números que o teto de páginas foi calibrado.

function printHistoryReport(historyStats) {
  console.log("\n  📊 Top 15 ativos mais lentos no /history/pastItems:");
  console.log("  " + "─".repeat(95));
  console.log("  " + ["páginas", "itens", "ms", "status?", "cc?", "esgot?", "ativo"].map((s, i) =>
    [8, 6, 7, 8, 5, 7, 0][i] ? s.padEnd([8, 6, 7, 8, 5, 7, 30][i]) : s).join(""));
  console.log("  " + "─".repeat(95));
  const top = [...historyStats].sort((a, b) => b.ms - a.ms).slice(0, 15);
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
  const totalPages = historyStats.reduce((sum, s) => sum + s.pages, 0);
  const totalItems = historyStats.reduce((sum, s) => sum + s.items_seen, 0);
  const exhausted  = historyStats.filter((s) => s.exhausted).length;
  const semStatus  = historyStats.filter((s) => !s.found_status).length;
  const semCC      = historyStats.filter((s) => !s.found_cc).length;
  console.log("  " + "─".repeat(95));
  console.log(`  Resumo: ${historyStats.length} ativos | ${totalPages} páginas (~${(totalPages / historyStats.length).toFixed(1)}/ativo) | ${totalItems} itens lidos`);
  console.log(`          ${exhausted} esgotaram histórico | ${semStatus} sem status encontrado | ${semCC} sem CC encontrado`);
  console.log();
}

module.exports = { printHistoryReport };
