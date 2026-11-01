// Gera os dois relatórios com dados FICTÍCIOS, sem TOPdesk nem SharePoint.
// Serve para ver o resultado e para a imagem do README.
//
//   npm run exemplo

const path = require("path");
const { formatDate } = require("../src/util");
const { motivoSemCC } = require("../src/regras");
const { writeExcel, writeSemCCExcel } = require("../src/excel");

// Gerador pseudoaleatório com semente fixa: o exemplo sai igual em toda execução.
function seeded(seed) {
  let x = seed;
  return () => ((x = (x * 16807) % 2147483647) / 2147483647);
}
const rand = seeded(42);
const pick = (list) => list[Math.floor(rand() * list.length)];

const UFS     = ["SP", "SP", "SP", "RJ", "RJ", "MG", "PR", "BA"];
const CCS     = ["1101 - Obra Alfa", "1102 - Obra Beta", "1205 - Obra Gama", "2001 - Administrativo", "3010 - Obra Delta"];
const MODELOS = {
  Notebook: ["Notebook 14\" i5", "Notebook 15\" i7", "Ultrabook 13\""],
  Celular:  ["Smartphone A", "Smartphone B"],
  Tablet:   ["Tablet 10\""],
};

const now = Date.now();
const rows = [];
for (let i = 1; i <= 48; i++) {
  const kind   = i % 4 === 0 ? "Celular" : i % 7 === 0 ? "Tablet" : "Notebook";
  const uf     = pick(UFS);
  const prefix = kind === "Notebook" ? "NB" : "CL";
  const nome   = `${prefix}-${uf}-${kind === "Tablet" ? "TABLET-" : ""}${String(100 + i).padStart(4, "0")}`;
  const dias   = Math.floor(rand() * 200);
  const data   = new Date(now - dias * 86_400_000);

  rows.push({
    nome_dispositivo:   nome,
    modelo:             pick(MODELOS[kind]),
    tipo:               kind === "Notebook" ? "Notebook" : "Celular",
    patrimonio:         `PAT-${String(5000 + i)}`,
    usuario:            `Colaborador ${String(i).padStart(2, "0")}`,
    centro_custo:       i % 9 === 0 ? "" : pick(CCS),
    uf,
    departamento:       pick(["Engenharia", "Suprimentos", "Financeiro", "Obras"]),
    dias_sem_retorno:   dias,
    data_mudanca:       formatDate(data),
    data_nao_devolvido: "",
    num_chamado:        `I${data.getFullYear()}-${String(1000 + i)}`,
  });
}
rows.sort((a, b) => b.dias_sem_retorno - a.dias_sem_retorno);

const semCCRows = rows
  .filter((r) => !r.centro_custo)
  .map((r, i) => ({
    ...r,
    motivo: motivoSemCC(i % 2 ? { exhausted: true, pages: 37 } : { hit_max_pages: true, pages: 100 }, 100),
  }));

(async () => {
  const principal = path.resolve("exemplo_nao_devolvidos.xlsx");
  const semCC     = path.resolve("exemplo_nao_devolvidos_sem_cc.xlsx");
  await writeExcel(rows, principal);
  await writeSemCCExcel(semCCRows, semCC);
  console.log(`Gerados com ${rows.length} ativos fictícios:\n  ${principal}\n  ${semCC}`);
})();
