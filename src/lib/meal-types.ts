export type CellMark = "none" | "ok" | "ng";

export type ColumnData = {
  judgment: CellMark;
  count: number;
  memo: string;
};

export type MealStore = Record<string, Record<string, ColumnData[]>>;

export type DayRecord = {
  dateLabel: string;
  morning: CellMark;
  lunch: CellMark;
  dinner: CellMark;
  stapleCount: number;
  memo: string;
};

export const MEAL_COLUMNS = ["朝", "昼", "夜"] as const;

export function createEmptyData(): ColumnData[] {
  return [
    { judgment: "none", count: 0, memo: "" },
    { judgment: "none", count: 0, memo: "" },
    { judgment: "none", count: 0, memo: "" },
  ];
}

export function dateKeyLocal(d: Date): string {
  const y = d.getFullYear();
  const m = String(d.getMonth() + 1).padStart(2, "0");
  const day = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

export function normalizeColumnData(raw: unknown): ColumnData[] | null {
  if (!Array.isArray(raw) || raw.length !== 3) return null;
  const marks: CellMark[] = ["none", "ok", "ng"];
  return raw.map((cell) => {
    const c = cell as Partial<ColumnData>;
    const judgment = marks.includes(c.judgment as CellMark)
      ? (c.judgment as CellMark)
      : "none";
    const count =
      typeof c.count === "number" && c.count >= 0 && c.count <= 99 ? c.count : 0;
    const memo = typeof c.memo === "string" ? c.memo : "";
    return { judgment, count, memo };
  });
}

export function columnDataToDayRecord(dateLabel: string, cols: ColumnData[]): DayRecord {
  const memoParts = cols
    .map((c, i) => {
      const t = c.memo.trim();
      return t ? `${MEAL_COLUMNS[i]}: ${t}` : "";
    })
    .filter(Boolean);
  return {
    dateLabel,
    morning: cols[0]?.judgment ?? "none",
    lunch: cols[1]?.judgment ?? "none",
    dinner: cols[2]?.judgment ?? "none",
    stapleCount: cols.reduce((s, c) => s + (c.count ?? 0), 0),
    memo: memoParts.join("\n"),
  };
}
