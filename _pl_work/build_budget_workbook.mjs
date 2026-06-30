import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const baseDir = path.dirname(fileURLToPath(import.meta.url));
const inputPath = path.join(baseDir, "budget_seed.json");
const outputPath = "C:\\Users\\PauloMendonça\\OneDrive - Redefrete\\Área de Trabalho\\Balanço\\DashBoard_P&L\\Budget.xlsx";

const data = JSON.parse(await fs.readFile(inputPath, "utf8"));

const theme = {
  navy: "#061B49",
  blue: "#1763E8",
  green: "#009E65",
  red: "#EF3E42",
  teal: "#008EAD",
  violet: "#7C3AED",
  header: "#F3F6FB",
  line: "#D9E2F1",
  muted: "#5B6B8A",
};

function aoaFromObjects(rows, columns) {
  return rows.map((row) => columns.map((col) => row[col.key] ?? ""));
}

function setHeader(range) {
  range.format = {
    fill: theme.header,
    font: { bold: true, color: theme.muted },
    borders: { preset: "all", style: "thin", color: theme.line },
  };
}

function styleTable(sheet, rangeAddress, currencyCols = []) {
  const range = sheet.getRange(rangeAddress);
  range.format = {
    borders: {
      insideHorizontal: { style: "thin", color: theme.line },
      top: { style: "thin", color: theme.line },
      bottom: { style: "thin", color: theme.line },
    },
  };
  const header = sheet.getRange(rangeAddress.replace(/\d+:.+/, "1:" + rangeAddress.split(":")[1].replace(/\d+/, "1")));
  setHeader(header);
  for (const col of currencyCols) {
    sheet.getRange(`${col}:${col}`).format.numberFormat = '"R$" #,##0.00;-"R$" #,##0.00';
  }
}

function safeSheetName(name) {
  return name.slice(0, 31);
}

function writeSheet(workbook, name, columns, rows, currencyCols = [], freezeRows = 1) {
  const sheet = workbook.worksheets.add(safeSheetName(name));
  sheet.showGridLines = false;
  const matrix = [columns.map((col) => col.label), ...aoaFromObjects(rows, columns)];
  sheet.getRangeByIndexes(0, 0, matrix.length, columns.length).values = matrix;
  setHeader(sheet.getRangeByIndexes(0, 0, 1, columns.length));
  sheet.freezePanes.freezeRows(freezeRows);
  const lastCol = String.fromCharCode(64 + Math.min(columns.length, 26));
  if (columns.length <= 26) {
    styleTable(sheet, `A1:${lastCol}${matrix.length}`, currencyCols);
  }
  columns.forEach((col, idx) => {
    const letter = String.fromCharCode(65 + idx);
    const width = col.width ?? 14;
    sheet.getRange(`${letter}:${letter}`).format.columnWidth = width;
    if (col.format) sheet.getRange(`${letter}:${letter}`).format.numberFormat = col.format;
  });
  return sheet;
}

const workbook = Workbook.create();

const resumo = workbook.worksheets.add("Resumo");
resumo.showGridLines = false;
resumo.getRange("A1:H1").merge();
resumo.getRange("A1").values = [["Budget automático - Média dos últimos 4 meses"]];
resumo.getRange("A1").format = {
  fill: theme.navy,
  font: { bold: true, color: "#FFFFFF", size: 16 },
};
resumo.getRange("A3:B11").values = [
  ["Data de geração", data.meta.generatedAt],
  ["Fonte", data.meta.source],
  ["Meses base", data.meta.basePeriodLabels.join(", ")],
  ["Meses projetados", data.meta.projectPeriodLabels.join(", ")],
  ["Combinações base", data.meta.baseRows],
  ["Linhas de budget", data.meta.budgetRows],
  ["Linhas de histórico", data.meta.historyRows],
  ["Cenário", "Budget Média 4M"],
  ["Observação", "Primeira versão sem ajuste de sazonalidade/campanhas."],
];
setHeader(resumo.getRange("A3:A11"));
resumo.getRange("A:B").format.columnWidth = 34;
resumo.getRange("B:B").format.columnWidth = 90;
resumo.getRange("B5:B7").format.numberFormat = "#,##0";

const budgetCols = [
  { key: "scenario", label: "Cenário", width: 18 },
  { key: "year", label: "Ano", width: 10, format: "0" },
  { key: "month", label: "Mês", width: 9, format: "0" },
  { key: "period", label: "Período", width: 12 },
  { key: "client", label: "Cliente", width: 18 },
  { key: "project", label: "Projeto", width: 22 },
  { key: "hub", label: "Unidade", width: 20 },
  { key: "expt", label: "Expt", width: 22 },
  { key: "vehicleType", label: "Tipo", width: 14 },
  { key: "fleetType", label: "Frota", width: 12 },
  { key: "account", label: "Conta DRE", width: 30 },
  { key: "category", label: "Categoria", width: 34 },
  { key: "costType", label: "Tipo Custo", width: 14 },
  { key: "budgetValue", label: "Valor Budget", width: 16, format: '"R$" #,##0.00;-"R$" #,##0.00' },
  { key: "note", label: "Observação", width: 34 },
];
writeSheet(workbook, "Budget", budgetCols, data.budgetRows, ["N"]);

const baseCols = [
  { key: "client", label: "Cliente", width: 18 },
  { key: "project", label: "Projeto", width: 22 },
  { key: "hub", label: "Unidade", width: 20 },
  { key: "expt", label: "Expt", width: 22 },
  { key: "vehicleType", label: "Tipo", width: 14 },
  { key: "fleetType", label: "Frota", width: 12 },
  { key: "account", label: "Conta DRE", width: 30 },
  { key: "category", label: "Categoria", width: 34 },
  { key: "costType", label: "Tipo Custo", width: 14 },
  ...data.meta.basePeriods.map((period, idx) => ({
    key: period,
    label: data.meta.basePeriodLabels[idx],
    width: 14,
    format: '"R$" #,##0.00;-"R$" #,##0.00',
  })),
  { key: "average", label: "Média 4M", width: 15, format: '"R$" #,##0.00;-"R$" #,##0.00' },
  { key: "activeMonths", label: "Meses c/ Movimento", width: 18, format: "0" },
];
writeSheet(workbook, "Base_Media_4M", baseCols, data.baseRows, ["J", "K", "L", "M", "N"]);

const histCols = [
  { key: "period", label: "Período", width: 12 },
  { key: "client", label: "Cliente", width: 18 },
  { key: "project", label: "Projeto", width: 22 },
  { key: "hub", label: "Unidade", width: 20 },
  { key: "expt", label: "Expt", width: 22 },
  { key: "vehicleType", label: "Tipo", width: 14 },
  { key: "fleetType", label: "Frota", width: 12 },
  { key: "account", label: "Conta DRE", width: 30 },
  { key: "category", label: "Categoria", width: 34 },
  { key: "costType", label: "Tipo Custo", width: 14 },
  { key: "actualValue", label: "Valor Realizado", width: 16, format: '"R$" #,##0.00;-"R$" #,##0.00' },
];
writeSheet(workbook, "Historico_Realizado", histCols, data.historyRows, ["K"]);

const campanhasRows = [
  {
    year: 2026,
    month: "",
    client: "",
    project: "",
    hub: "",
    campaign: "",
    seasonalFactor: "",
    note: "Preencher depois para ajustar o budget por campanha/sazonalidade.",
  },
];
writeSheet(
  workbook,
  "Campanhas",
  [
    { key: "year", label: "Ano", width: 10 },
    { key: "month", label: "Mês", width: 10 },
    { key: "client", label: "Cliente", width: 18 },
    { key: "project", label: "Projeto", width: 22 },
    { key: "hub", label: "Unidade", width: 20 },
    { key: "campaign", label: "Campanha", width: 26 },
    { key: "seasonalFactor", label: "Fator Sazonal", width: 16, format: "0.0%" },
    { key: "note", label: "Observação", width: 54 },
  ],
  campanhasRows,
);

const premissasRows = [
  {
    metric: "Base do Budget",
    dimension: "Geral",
    value: "Média dos últimos 4 meses realizados",
    start: data.meta.projectPeriodLabels[0],
    end: data.meta.projectPeriodLabels.at(-1),
    note: "Gerado automaticamente; revisar sazonalidade e eventos pontuais.",
  },
];
writeSheet(
  workbook,
  "Premissas",
  [
    { key: "metric", label: "Indicador", width: 24 },
    { key: "dimension", label: "Dimensão", width: 20 },
    { key: "value", label: "Valor / Regra", width: 44 },
    { key: "start", label: "Vigência Inicial", width: 16 },
    { key: "end", label: "Vigência Final", width: 16 },
    { key: "note", label: "Observação", width: 54 },
  ],
  premissasRows,
);

const inspect = await workbook.inspect({
  kind: "sheet,region",
  sheetId: "Resumo",
  range: "A1:B11",
  maxChars: 2000,
});
console.log(inspect.ndjson);

const errors = await workbook.inspect({
  kind: "match",
  searchTerm: "#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A",
  options: { useRegex: true, maxResults: 50 },
  maxChars: 2000,
});
console.log(errors.ndjson);

const preview = await workbook.render({ sheetName: "Resumo", autoCrop: "all", scale: 1, format: "png" });
await fs.writeFile(path.join(baseDir, "budget_resumo_preview.png"), new Uint8Array(await preview.arrayBuffer()));

await fs.mkdir(path.dirname(outputPath), { recursive: true });
const xlsx = await SpreadsheetFile.exportXlsx(workbook);
await xlsx.save(outputPath);
console.log(outputPath);
