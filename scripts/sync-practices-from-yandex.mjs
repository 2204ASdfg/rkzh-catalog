import fs from "node:fs/promises";
import * as XLSX from "xlsx";

const publicUrl = process.env.YANDEX_PUBLIC_URL;
const outputPath = "site/data.js";

if (!publicUrl) {
  throw new Error("Не задана переменная YANDEX_PUBLIC_URL.");
}

const metadataUrl = new URL("https://cloud-api.yandex.net/v1/disk/public/resources/download");
metadataUrl.searchParams.set("public_key", publicUrl);

const metadataResponse = await fetch(metadataUrl, {
  headers: { Accept: "application/json" },
});

if (!metadataResponse.ok) {
  throw new Error(`Яндекс Диск не отдал ссылку на скачивание: HTTP ${metadataResponse.status}.`);
}

const { href } = await metadataResponse.json();

if (!href) {
  throw new Error("Яндекс Диск не вернул адрес для скачивания Excel-файла.");
}

const fileResponse = await fetch(href);

if (!fileResponse.ok) {
  throw new Error(`Не удалось скачать Excel-файл: HTTP ${fileResponse.status}.`);
}

const workbook = XLSX.read(Buffer.from(await fileResponse.arrayBuffer()), { type: "buffer" });
const sheetName = workbook.SheetNames.find((name) => name === "Лучшие практики") || workbook.SheetNames[0];

if (!sheetName) {
  throw new Error("В Excel-файле не найден лист с практиками.");
}

const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], {
  header: 1,
  defval: "",
  raw: false,
});

const headerRowIndex = rows.findIndex((row) => row.some((cell) => String(cell).trim() === "№"));

if (headerRowIndex < 0) {
  throw new Error("В Excel-файле не найдена строка заголовков с колонкой «№».");
}

const headers = rows[headerRowIndex].map((cell) => String(cell).trim());
const practices = rows
  .slice(headerRowIndex + 1)
  .filter((row) => String(row[0] ?? "").trim())
  .map((row) => Object.fromEntries(headers.map((header, index) => [header, String(row[index] ?? "").trim()])))
  .map((row) => ({
    id: row["№"],
    name: row["Название практики"],
    country: row["Страна"],
    region: row["Регион"],
    direction: row["Направление РКЖ"],
    factor: row["Фактор РКЖ"],
    indicator: row["Показатель РКЖ"],
    audience: row["Целевая аудитория"],
    problem: row["Проблема"],
    essence: row["Суть практики"],
    success: row["Ключевые факторы успеха"],
    source: row["Источник информации"],
    years: row["Год(ы) реализации"],
    budget: row["Оценка стоимости / бюджет реализации"],
    international: String(row["Страна"] || "").trim().toLocaleLowerCase("ru") !== "россия",
  }));

if (!practices.length) {
  throw new Error("После чтения Excel-файла не найдено ни одной практики.");
}

await fs.writeFile(outputPath, `window.PRACTICES = ${JSON.stringify(practices, null, 2)};\n`, "utf8");
console.log(`Готово: ${practices.length} практик записано в ${outputPath}.`);
