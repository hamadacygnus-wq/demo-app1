import {
  normalizeColumnData,
  type CellMark,
  type ColumnData,
  type MealStore,
} from "./meal-types";
import { resolveWorkbookTarget } from "./excel-resolve";
import { getGraphAccessToken, graphFetch, isGraphConfigured } from "./graph-auth";

/** Excel 1行目のヘッダー（この順序で作成してください） */
const HEADERS = [
  "利用者",
  "日付",
  "朝判定",
  "昼判定",
  "夜判定",
  "朝カウント",
  "昼カウント",
  "夜カウント",
  "朝メモ",
  "昼メモ",
  "夜メモ",
] as const;

function worksheetName(): string {
  return process.env.GRAPH_WORKSHEET_NAME?.trim() || "記録";
}

async function getTarget() {
  return resolveWorkbookTarget();
}

function workbookBase(driveId: string, itemId: string): string {
  return `/drives/${driveId}/items/${itemId}/workbook/worksheets('${encodeURIComponent(worksheetName())}')`;
}

function parseJudgment(value: unknown): CellMark {
  const s = String(value ?? "").trim().toLowerCase();
  if (s === "ok" || s === "○" || s === "◯" || s === "maru") return "ok";
  if (s === "ng" || s === "×" || s === "batsu" || s === "x") return "ng";
  return "none";
}

function parseCount(value: unknown): number {
  const n = Number(value);
  if (!Number.isFinite(n) || n < 0) return 0;
  return Math.min(99, Math.floor(n));
}

function rowToColumnData(row: unknown[]): ColumnData[] | null {
  if (row.length < 11) return null;
  const cols: ColumnData[] = [
    {
      judgment: parseJudgment(row[2]),
      count: parseCount(row[5]),
      memo: String(row[8] ?? ""),
    },
    {
      judgment: parseJudgment(row[3]),
      count: parseCount(row[6]),
      memo: String(row[9] ?? ""),
    },
    {
      judgment: parseJudgment(row[4]),
      count: parseCount(row[7]),
      memo: String(row[10] ?? ""),
    },
  ];
  return normalizeColumnData(cols);
}

function columnDataToRow(user: string, date: string, data: ColumnData[]): (string | number)[] {
  return [
    user,
    date,
    data[0]?.judgment ?? "none",
    data[1]?.judgment ?? "none",
    data[2]?.judgment ?? "none",
    data[0]?.count ?? 0,
    data[1]?.count ?? 0,
    data[2]?.count ?? 0,
    data[0]?.memo ?? "",
    data[1]?.memo ?? "",
    data[2]?.memo ?? "",
  ];
}

function rowsToMealStore(rows: unknown[][]): MealStore {
  const store: MealStore = {};
  for (const row of rows) {
    if (!Array.isArray(row) || row.length < 2) continue;
    const user = String(row[0] ?? "").trim();
    const date = String(row[1] ?? "").trim();
    if (!user || !date || user === HEADERS[0]) continue;
    const cols = rowToColumnData(row);
    if (!cols) continue;
    if (!store[user]) store[user] = {};
    store[user][date] = cols;
  }
  return store;
}

function mealStoreToRows(store: MealStore): (string | number)[][] {
  const rows: (string | number)[][] = [HEADERS as unknown as string[]];
  const keys = Object.keys(store).sort();
  for (const user of keys) {
    const dates = Object.keys(store[user] ?? {}).sort();
    for (const date of dates) {
      const data = store[user][date];
      if (!data) continue;
      rows.push(columnDataToRow(user, date, data));
    }
  }
  return rows;
}

function columnLetter(col: number): string {
  let n = col;
  let s = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    s = String.fromCharCode(65 + rem) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

function rangeAddress(rowCount: number, colCount: number): string {
  const endCol = columnLetter(colCount);
  return `A1:${endCol}${rowCount}`;
}

export async function readMealStoreFromExcel(): Promise<MealStore> {
  const { driveId, itemId } = await getTarget();
  const base = workbookBase(driveId, itemId);

  const res = await graphFetch(`${base}/usedRange`);
  if (res.status === 404) {
    return {};
  }
  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Excelの読み取りに失敗しました: ${res.status} ${text}`);
  }

  const body = (await res.json()) as { values?: unknown[][] };
  const values = body.values ?? [];
  if (values.length === 0) {
    return {};
  }

  const firstCell = String(values[0]?.[0] ?? "").trim();
  const dataRows =
    firstCell === HEADERS[0] ? values.slice(1) : values;

  return rowsToMealStore(dataRows);
}

async function ensureWorksheetWithHeaders(token: string, driveId: string, itemId: string) {
  const sheet = worksheetName();
  const listRes = await graphFetch(`/drives/${driveId}/items/${itemId}/workbook/worksheets`);
  if (!listRes.ok) {
    throw new Error(`ワークシート一覧の取得に失敗しました: ${listRes.status}`);
  }
  const list = (await listRes.json()) as { value?: { name?: string }[] };
  const exists = list.value?.some((w) => w.name === sheet);
  if (!exists) {
    const addRes = await graphFetch(
      `/drives/${driveId}/items/${itemId}/workbook/worksheets/add`,
      {
        method: "POST",
        body: JSON.stringify({ name: sheet }),
      },
    );
    if (!addRes.ok) {
      const text = await addRes.text();
      throw new Error(`ワークシートの作成に失敗しました: ${addRes.status} ${text}`);
    }
  }

  const base = workbookBase(driveId, itemId);
  const headerRange = `${base}/range(address='A1:K1')`;
  await fetch(`https://graph.microsoft.com/v1.0${headerRange}`, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ values: [HEADERS] }),
  });
}

export async function writeMealStoreToExcel(store: MealStore): Promise<void> {
  const token = await getGraphAccessToken();
  const { driveId, itemId } = await getTarget();
  await ensureWorksheetWithHeaders(token, driveId, itemId);

  const rows = mealStoreToRows(store);
  const address = rangeAddress(rows.length, HEADERS.length);
  const base = workbookBase(driveId, itemId);
  const rangeUrl = `https://graph.microsoft.com/v1.0${base}/range(address='${address}')`;

  const res = await fetch(rangeUrl, {
    method: "PATCH",
    headers: {
      Authorization: `Bearer ${token}`,
      "Content-Type": "application/json",
    },
    body: JSON.stringify({ values: rows }),
  });

  if (!res.ok) {
    const text = await res.text();
    throw new Error(`Excelへの書き込みに失敗しました: ${res.status} ${text}`);
  }
}

export async function upsertMealRecord(
  user: string,
  date: string,
  data: ColumnData[],
): Promise<void> {
  const store = await readMealStoreFromExcel();
  if (!store[user]) store[user] = {};
  store[user][date] = data.map((c) => ({ ...c }));
  await writeMealStoreToExcel(store);
}

export function mergeMealStores(base: MealStore, patch: MealStore): MealStore {
  const merged: MealStore = { ...base };
  for (const user of Object.keys(patch)) {
    merged[user] = { ...(merged[user] ?? {}), ...patch[user] };
  }
  return merged;
}

export { isGraphConfigured };
