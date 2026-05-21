import { NextResponse } from "next/server";
import { readMealStoreFromExcel, upsertMealRecord } from "@/lib/excel-meals";
import { isGraphConfigured } from "@/lib/graph-auth";
import { normalizeColumnData, type ColumnData } from "@/lib/meal-types";

export async function GET() {
  if (!isGraphConfigured()) {
    return NextResponse.json(
      {
        configured: false,
        store: {},
        message: "Microsoft Graph の環境変数が未設定のため、Excel からは取得していません。",
      },
      { status: 200 },
    );
  }

  try {
    const store = await readMealStoreFromExcel();
    return NextResponse.json({ configured: true, store, source: "excel" as const });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Excel の読み取りに失敗しました";
    console.error("[GET /api/meals]", error);
    return NextResponse.json({ configured: true, error: message }, { status: 500 });
  }
}

export async function POST(request: Request) {
  if (!isGraphConfigured()) {
    return NextResponse.json(
      { error: "Microsoft Graph の環境変数が未設定です。.env.local を確認してください。" },
      { status: 503 },
    );
  }

  let body: unknown;
  try {
    body = await request.json();
  } catch {
    return NextResponse.json({ error: "リクエストボディが不正です" }, { status: 400 });
  }

  const { user, date, data } = body as {
    user?: string;
    date?: string;
    data?: ColumnData[];
  };

  if (!user?.trim() || !date?.trim()) {
    return NextResponse.json({ error: "利用者と日付は必須です" }, { status: 400 });
  }

  const normalized = normalizeColumnData(data);
  if (!normalized) {
    return NextResponse.json({ error: "食事データの形式が不正です" }, { status: 400 });
  }

  try {
    await upsertMealRecord(user.trim(), date.trim(), normalized);
    return NextResponse.json({ ok: true });
  } catch (error) {
    const message = error instanceof Error ? error.message : "Excel への書き込みに失敗しました";
    console.error("[POST /api/meals]", error);
    return NextResponse.json({ error: message }, { status: 500 });
  }
}
