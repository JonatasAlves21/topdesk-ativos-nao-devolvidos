// Aba "Resumo": KPIs, faixas de prazo, totais por UF e por centro de custo.

const { PAL_DARK, PAL_ORANGE, PAL_ALT, PAL_WHITE, PAL_TEXT, AGING } = require("./estilo");

// Tablets são cadastrados no template de celular; o nome do dispositivo diferencia.
function isTablet(r) {
  return r.tipo === "Celular" && (r.nome_dispositivo || "").toLowerCase().includes("tablet");
}

// Números da aba Resumo, separados do desenho da planilha para poderem ser testados.
function summarize(rows) {
  const total     = rows.length;
  const notebooks = rows.filter((r) => r.tipo === "Notebook").length;
  const tablets   = rows.filter(isTablet).length;
  const celulares = rows.filter((r) => r.tipo === "Celular" && !isTablet(r)).length;

  const agingStats = AGING.map((a, i) => ({
    label: a.label,
    count: rows.filter((r) => {
      if (typeof r.dias_sem_retorno !== "number") return false;
      const min = i > 0 ? AGING[i - 1].max : -1;
      return r.dias_sem_retorno > min && r.dias_sem_retorno <= a.max;
    }).length,
    bg:   a.bg,
    text: a.text,
  }));
  const semData = rows.filter((r) => typeof r.dias_sem_retorno !== "number").length;

  const byUF = {};
  for (const r of rows) byUF[r.uf || "Outros"] = (byUF[r.uf || "Outros"] || 0) + 1;
  const ufEntries = Object.entries(byUF).sort((a, b) => b[1] - a[1]);

  // Pivot: CC → { notebook, celular, tablet, total }
  const byCC = {};
  for (const r of rows) {
    const cc  = r.centro_custo || "(Não informado)";
    if (!byCC[cc]) byCC[cc] = { notebook: 0, celular: 0, tablet: 0 };
    if (r.tipo === "Notebook")                          byCC[cc].notebook++;
    else if (isTablet(r))                               byCC[cc].tablet++;
    else                                                byCC[cc].celular++;
  }
  const ccEntries = Object.entries(byCC)
    .map(([cc, v]) => [cc, v.notebook, v.celular, v.tablet, v.notebook + v.celular + v.tablet])
    .sort((a, b) => b[4] - a[4]); // maior total primeiro

  return { total, notebooks, tablets, celulares, agingStats, semData, ufEntries, ccEntries };
}

function buildSummarySheet(wb, rows) {
  const ws = wb.addWorksheet("Resumo");
  const { total, notebooks, tablets, celulares, agingStats, semData, ufEntries, ccEntries } = summarize(rows);

  // helpers
  const gc = (addr) => ws.getCell(addr);
  const gr = (n)    => ws.getRow(n);

  function styleCell(addr, { value, bold, size, color, bg, halign, valign, wrapText, borderBottom, borderLeft } = {}) {
    const c = gc(addr);
    if (value !== undefined) c.value = value;
    c.font      = { bold: !!bold, size: size || 10, color: { argb: color || PAL_TEXT }, name: "Calibri" };
    c.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: bg || PAL_WHITE } };
    c.alignment = { horizontal: halign || "left", vertical: valign || "middle", wrapText: !!wrapText };
    const border = {};
    if (borderBottom) border.bottom = { style: borderBottom, color: { argb: PAL_ORANGE } };
    if (borderLeft)   border.left   = { style: borderLeft,   color: { argb: PAL_ORANGE } };
    c.border = border;
  }

  function sectionTitle(addr, text, mergeEnd) {
    if (mergeEnd) ws.mergeCells(`${addr}:${mergeEnd}`);
    styleCell(addr, {
      value: `  ${text}`, bold: true, size: 11,
      color: PAL_WHITE, bg: PAL_DARK, halign: "left",
      borderBottom: "medium", borderLeft: "medium",
    });
  }

  function tableHeader(rowN, cols) {
    cols.forEach((text, i) => {
      const c = gr(rowN).getCell(i + 1);
      c.value     = text;
      c.font      = { bold: true, color: { argb: PAL_WHITE }, name: "Calibri", size: 10 };
      c.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: PAL_DARK } };
      c.alignment = { horizontal: "center", vertical: "middle" };
      c.border    = { bottom: { style: "medium", color: { argb: PAL_ORANGE } } };
    });
    gr(rowN).height = 24;
  }

  function tableRow(rowN, cols, isAlt, overrideBg, overrideText) {
    cols.forEach((val, i) => {
      const c = gr(rowN).getCell(i + 1);
      c.value     = val;
      c.font      = { color: { argb: overrideText || PAL_TEXT }, name: "Calibri", size: 10 };
      c.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: overrideBg || (isAlt ? PAL_ALT : PAL_WHITE) } };
      c.alignment = { horizontal: i === 0 ? "left" : "center", vertical: "middle" };
      c.border    = { bottom: { style: "thin", color: { argb: "FFE0E0E0" } } };
    });
    gr(rowN).height = 20;
  }

  // ── Colunas ──────────────────────────────────────────────────────────────────
  [26, 13, 13, 10, 10, 12].forEach((w, i) => { ws.getColumn(i + 1).width = w; });

  // ── Título ───────────────────────────────────────────────────────────────────
  ws.mergeCells("A1:F1");
  styleCell("A1", {
    value: "NÃO DEVOLVIDOS — RESUMO EXECUTIVO",
    bold: true, size: 14, color: PAL_WHITE, bg: PAL_DARK, halign: "center",
    borderBottom: "thick",
  });
  gr(1).height = 44;

  ws.mergeCells("A2:F2");
  styleCell("A2", {
    value: `Atualizado em ${new Date().toLocaleDateString("pt-BR")}`,
    size: 9, color: "FF999999", bg: PAL_WHITE, halign: "center",
  });
  gr(2).height = 18;

  gr(3).height = 8; // espaçador

  // ── KPI boxes ────────────────────────────────────────────────────────────────
  // Labels (row 4)
  ws.mergeCells("A4:B4"); styleCell("A4", { value: "TOTAL GERAL",         bold: true, size: 9, color: "FF666666", bg: PAL_ALT, halign: "center" });
  ws.mergeCells("C4:D4"); styleCell("C4", { value: "NOTEBOOKS",           bold: true, size: 9, color: "FF666666", bg: PAL_ALT, halign: "center" });
  ws.mergeCells("E4:F4"); styleCell("E4", { value: "CELULARES / TABLETS", bold: true, size: 9, color: "FF666666", bg: PAL_ALT, halign: "center" });
  gr(4).height = 22;

  // Values (row 5)
  ws.mergeCells("A5:B5"); styleCell("A5", { value: total,               bold: true, size: 22, color: PAL_ORANGE, bg: PAL_WHITE, halign: "center" });
  ws.mergeCells("C5:D5"); styleCell("C5", { value: notebooks,           bold: true, size: 22, color: PAL_DARK,   bg: PAL_WHITE, halign: "center" });
  ws.mergeCells("E5:F5"); styleCell("E5", { value: `${celulares} / ${tablets}`, bold: true, size: 22, color: PAL_DARK, bg: PAL_WHITE, halign: "center" });
  gr(5).height = 48;

  // Bordas inferiores dos KPIs
  ["A4","C4","E4"].forEach((a) => { gc(a).border = { ...gc(a).border, top: { style: "thin", color: { argb: "FFD0D0D0" } }, left: { style: "thin", color: { argb: "FFD0D0D0" } }, right: { style: "thin", color: { argb: "FFD0D0D0" } } }; });
  ["A5","C5","E5"].forEach((a) => { gc(a).border = { bottom: { style: "thin", color: { argb: "FFD0D0D0" } }, left: { style: "thin", color: { argb: "FFD0D0D0" } }, right: { style: "thin", color: { argb: "FFD0D0D0" } } }; });

  gr(6).height = 10; // espaçador

  // ── Prazo sem retorno ────────────────────────────────────────────────────────
  const pct = (n) => total > 0 ? `${Math.round((n / total) * 100)}%` : "0%";

  sectionTitle("A7", "PRAZO SEM RETORNO", "F7");
  gr(7).height = 28;
  tableHeader(8, ["FAIXA", "QUANTIDADE", "% DO TOTAL"]);

  agingStats.forEach(({ label, count, bg, text }, i) => {
    tableRow(9 + i, [label, count, pct(count)], false, bg, text);
  });
  if (semData > 0) {
    const r = 9 + agingStats.length;
    tableRow(r, ["Sem data registrada", semData, pct(semData)], false, PAL_ALT, PAL_TEXT);
  }

  let nextRow = 9 + agingStats.length + (semData > 0 ? 1 : 0) + 2;

  // ── Por UF ───────────────────────────────────────────────────────────────────
  sectionTitle(`A${nextRow}`, "POR ESTADO (UF)", `C${nextRow}`);
  gr(nextRow).height = 28;
  tableHeader(nextRow + 1, ["UF", "QUANTIDADE", "% DO TOTAL"]);
  ufEntries.forEach(([uf, count], i) => {
    tableRow(nextRow + 2 + i, [uf, count, pct(count)], i % 2 === 1);
  });
  nextRow += 2 + ufEntries.length + 2;

  // ── Por Centro de Custo × Tipo de device ─────────────────────────────────────
  sectionTitle(`A${nextRow}`, "POR CENTRO DE CUSTO", `F${nextRow}`);
  gr(nextRow).height = 28;
  tableHeader(nextRow + 1, ["CENTRO DE CUSTO", "NOTEBOOKS", "CELULARES", "TABLETS", "TOTAL", "% DO TOTAL"]);
  ccEntries.forEach(([cc, nb, cel, tab, tot], i) => {
    tableRow(nextRow + 2 + i, [cc, nb || "-", cel || "-", tab || "-", tot, pct(tot)], i % 2 === 1);
  });

  // ── Nota de escopo ───────────────────────────────────────────────────────────
  // Evita "chamado de divergência": este consolidado e os arquivos diários por
  // obra (fluxo Power Automate) têm regras diferentes de propósito.
  const notaRow = nextRow + 2 + ccEntries.length + 1;
  ws.mergeCells(`A${notaRow}:F${notaRow}`);
  styleCell(`A${notaRow}`, {
    value:
      "Nota de escopo: este consolidado inclui ativos SEM centro de custo (detalhados em Nao_Devolvidos_Sem_CC.xlsx) e considera " +
      "mudanças de status desde 01/2026. Os arquivos diários ATIVOS_NAO_DEVOLVIDOS_OBRA_XXXX (por obra) omitem ativos sem CC e " +
      "não filtram por data — por isso os totais podem diferir entre os relatórios.",
    size: 8, color: "FF999999", halign: "left", wrapText: true,
  });
  gr(notaRow).height = 34;

  ws.views = [{ state: "frozen", ySplit: 1 }];
}

module.exports = { isTablet, summarize, buildSummarySheet };
