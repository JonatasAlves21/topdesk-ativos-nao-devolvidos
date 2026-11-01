// ── Paleta do relatório ───────────────────────────────────────────────────────
const PAL_DARK   = "FF333333"; // cinza escuro — cabeçalhos
const PAL_ORANGE = "FFF57C20"; // laranja de destaque
const PAL_ALT    = "FFF5F5F5"; // cinza claríssimo — linhas alternadas
const PAL_WHITE  = "FFFFFFFF";
const PAL_TEXT   = "FF1A1A1A"; // texto de dados

// Aging (faixa de prazo sem retorno)
const AGING = [
  { max: 30,  bg: "FFE2EFDA", text: "FF375623", label: "≤ 30 dias  (Em dia)"         },
  { max: 90,  bg: "FFFFEB9C", text: "FF9C5700", label: "31 a 90 dias  (Atenção)"     },
  { max: Infinity, bg: "FFFFC7CE", text: "FF9C0006", label: "> 90 dias  (Crítico)"   },
];

function agingBucket(dias) {
  if (typeof dias !== "number") return null;
  return AGING.find((a) => dias <= a.max) || AGING[AGING.length - 1];
}

module.exports = { PAL_DARK, PAL_ORANGE, PAL_ALT, PAL_WHITE, PAL_TEXT, AGING, agingBucket };
