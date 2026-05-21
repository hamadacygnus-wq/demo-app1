"use client";

import {
  CalendarDays,
  ChevronLeft,
  ChevronRight,
  Circle,
  Eraser,
  X,
  UtensilsCrossed,
  Home,
} from "lucide-react";
import { useCallback, useEffect, useState } from "react";
import { loadMealStoreLocal, saveMealStoreLocal } from "@/lib/meal-store-local";
import {
  columnDataToDayRecord,
  createEmptyData,
  dateKeyLocal,
  MEAL_COLUMNS,
  normalizeColumnData,
  type CellMark,
  type ColumnData,
  type DayRecord,
  type MealStore,
} from "@/lib/meal-types";

export type AppView = "HOME" | "INPUT" | "HISTORY";

type MarkTool = "ok" | "ng" | "clear";

/** 青色の一重丸（○） */
const OK_MARK_CLASS = {
  large: "h-12 w-12",
  medium: "h-11 w-11",
  small: "h-6 w-6 mx-auto",
} as const;

function OkMark({ size }: { size: keyof typeof OK_MARK_CLASS }) {
  return (
    <Circle
      className={`${OK_MARK_CLASS[size]} text-blue-500`}
      strokeWidth={2}
      fill="none"
    />
  );
}

const NG_MARK_CLASS = {
  large: "h-12 w-12",
  medium: "h-11 w-11",
  small: "h-6 w-6 mx-auto",
} as const;

function NgMark({ size }: { size: keyof typeof NG_MARK_CLASS }) {
  return <X className={`${NG_MARK_CLASS[size]} text-red-500`} strokeWidth={2} />;
}

export function MealMainScreen() {
  const [view, setView] = useState<AppView>("HOME");
  const [selectedUser, setSelectedUser] = useState("");
  const [currentDate, setCurrentDate] = useState<Date>(new Date());
  const [activeTool, setActiveTool] = useState<MarkTool>("ok");
  const [tableData, setTableData] = useState<ColumnData[]>(createEmptyData());
  const [mealStore, setMealStore] = useState<MealStore>(() => loadMealStoreLocal());
  const [isSyncing, setIsSyncing] = useState(true);
  const [excelConfigured, setExcelConfigured] = useState(false);
  const [syncError, setSyncError] = useState<string | null>(null);
  const [isSending, setIsSending] = useState(false);

  const [users] = useState<string[]>([
    "A棟　シグナス太郎　様",
    "A棟　シグナス花子　様",
    "B棟　オリオン次郎　様",
    "C棟　リラ三郎　様",
  ]);

  const [monthlyHistory, setMonthlyHistory] = useState<DayRecord[]>([]);

  const applyMealStore = useCallback((store: MealStore) => {
    setMealStore(store);
    saveMealStoreLocal(store);
  }, []);

  useEffect(() => {
    let cancelled = false;
    async function syncFromExcel() {
      setIsSyncing(true);
      setSyncError(null);
      try {
        const res = await fetch("/api/meals");
        const json = (await res.json()) as {
          configured?: boolean;
          store?: MealStore;
          error?: string;
        };
        if (cancelled) return;
        setExcelConfigured(Boolean(json.configured));
        if (!json.configured) {
          setMealStore(loadMealStoreLocal());
          return;
        }
        if (!res.ok) {
          setSyncError(json.error ?? "Excel からの取得に失敗しました");
          setMealStore(loadMealStoreLocal());
          return;
        }
        if (json.store) {
          applyMealStore(json.store);
        }
      } catch {
        if (!cancelled) {
          setSyncError("Excel との通信に失敗しました。ブラウザ内の保存データを表示しています。");
          setMealStore(loadMealStoreLocal());
        }
      } finally {
        if (!cancelled) setIsSyncing(false);
      }
    }
    syncFromExcel();
    return () => { cancelled = true; };
  }, [applyMealStore]);

  const fetchDailyData = useCallback(
    (userName: string, date: Date) => {
      const key = dateKeyLocal(date);
      const normalized = normalizeColumnData(mealStore[userName]?.[key]);
      setTableData(normalized ?? createEmptyData());
    },
    [mealStore],
  );

  const fetchMonthlyHistory = useCallback(
    (userName: string, date: Date) => {
      const year = date.getFullYear();
      const month = date.getMonth();
      const lastDay = new Date(year, month + 1, 0).getDate();
      const weekdayLabels = ["日", "月", "火", "水", "木", "金", "土"] as const;

      const days: DayRecord[] = [];
      for (let d = 1; d <= lastDay; d++) {
        const target = new Date(year, month, d);
        const dayLabel = `${month + 1}/${d}(${weekdayLabels[target.getDay()]})`;
        const normalized = normalizeColumnData(
          mealStore[userName]?.[dateKeyLocal(target)],
        );
        if (normalized) {
          days.push(columnDataToDayRecord(dayLabel, normalized));
        } else {
          days.push({
            dateLabel: dayLabel,
            morning: "none", lunch: "none", dinner: "none",
            stapleCount: 0, memo: "",
          });
        }
      }
      setMonthlyHistory(days);
    },
    [mealStore],
  );

  const onSend = useCallback(async () => {
    if (window.confirm("この内容で送信しますか？")) {
      const dateKey = dateKeyLocal(currentDate);
      const snapshot = tableData.map((c) => ({ ...c }));

      const nextStore: MealStore = {
        ...mealStore,
        [selectedUser]: {
          ...(mealStore[selectedUser] ?? {}),
          [dateKey]: snapshot,
        },
      };
      applyMealStore(nextStore);

      if (!excelConfigured) {
        alert("送信完了");
        setView("HOME");
        setSelectedUser("");
        return;
      }

      setIsSending(true);
      try {
        const res = await fetch("/api/meals", {
          method: "POST",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({
            user: selectedUser,
            date: dateKey,
            data: snapshot,
          }),
        });
        const json = (await res.json()) as { error?: string };
        if (!res.ok) {
          throw new Error(json.error ?? "Excel への書き込みに失敗しました");
        }
        alert("送信完了");
        setView("HOME");
        setSelectedUser("");
      } catch (error) {
        const message = error instanceof Error ? error.message : "Excel への書き込みに失敗しました";
        alert(`送信に失敗しました。\n${message}`);
      } finally {
        setIsSending(false);
      }
    }
  }, [selectedUser, currentDate, tableData, mealStore, applyMealStore, excelConfigured]);

  useEffect(() => {
    if (selectedUser) {
      fetchDailyData(selectedUser, currentDate);
    }
  }, [currentDate, selectedUser, fetchDailyData]);

  const formattedDate = new Intl.DateTimeFormat("ja-JP", {
    month: "long", day: "numeric", weekday: "short",
  }).format(currentDate);

  const changeDate = (days: number) => {
    setCurrentDate((prev) => {
      const nextDate = new Date(prev);
      nextDate.setDate(nextDate.getDate() + days);
      return nextDate;
    });
  };

  const onJudgmentPress = (colIdx: number) => {
    setTableData((prev) => {
      const next = [...prev];
      if (activeTool === "clear") {
        // 消しゴムツールが選ばれている時は消去する
        next[colIdx] = { ...next[colIdx], judgment: "none" };
      } else {
        // 〇か×が選ばれている時は、常にそのマークで上書きする（再タップで消えない）
        const mark: CellMark = activeTool === "ok" ? "ok" : "ng";
        next[colIdx] = {
          ...next[colIdx],
          judgment: mark,
        };
      }
      return next;
    });
  };

  const onCountPress = (colIdx: number) => {
    setTableData((prev) => {
      const next = [...prev];
      if (activeTool === "clear") {
        next[colIdx] = { ...next[colIdx], count: 0 };
      } else {
        next[colIdx] = {
          ...next[colIdx],
          count: next[colIdx].count >= 99 ? 0 : next[colIdx].count + 1,
        };
      }
      return next;
    });
  };

  return (
    <div className="flex h-screen max-h-screen flex-col bg-slate-100 font-sans overflow-hidden select-none">
      {isSyncing && (
        <div className="bg-amber-100 py-1.5 text-center text-xs font-bold text-amber-900 shrink-0">
          Excel からデータを読み込んでいます…
        </div>
      )}
      {!isSyncing && syncError && (
        <div className="bg-red-100 py-1.5 text-center text-xs font-bold text-red-900 shrink-0">
          {syncError}
        </div>
      )}

      {/* HOME画面 */}
      {view === "HOME" && (
        <div className="flex flex-1 flex-col p-4 md:p-6 max-w-5xl w-full mx-auto justify-between overflow-hidden">
          <div className="flex flex-col items-center justify-center bg-[#e1f3fb] rounded-2xl p-6 md:p-10 shadow-sm border border-slate-200">
            <div className="flex items-center gap-3 mb-6 md:mb-8 bg-white/60 px-6 py-2.5 rounded-full shadow-inner border border-sky-100">
              <UtensilsCrossed className="h-7 w-7 text-sky-500 stroke-[2.5]" />
              <h1 className="text-2xl md:text-3xl font-black text-slate-800 tracking-wider">食事管理アプリ</h1>
            </div>
            <p className="mb-6 text-lg font-bold text-slate-600">利用者を選んでください。</p>
            
            <div className="grid w-full max-w-3xl grid-cols-2 gap-4 md:gap-6">
              {users.map((user) => (
                <button
                  key={user}
                  onClick={() => setSelectedUser(user)}
                  className={`rounded-xl py-5 md:py-7 text-center text-xl md:text-2xl font-bold shadow-sm transition active:scale-95 duration-100 ${
                    selectedUser === user
                      ? "bg-sky-500 text-white ring-4 ring-sky-200"
                      : "bg-white text-slate-800 hover:bg-sky-50 border border-slate-200"
                  }`}
                >
                  {user.replace("　", " ")}
                </button>
              ))}
            </div>
          </div>

          <div className="flex justify-center py-4 md:py-8 shrink-0">
            <button
              onClick={() => setView("INPUT")}
              disabled={!selectedUser || isSyncing}
              className={`rounded-2xl px-20 py-5 text-2xl font-bold tracking-wider shadow-md transition active:scale-95 duration-100 ${
                selectedUser && !isSyncing
                  ? "bg-[#a7def5] text-slate-800 hover:bg-[#8ccfed]"
                  : "bg-slate-200 text-slate-400 opacity-50 cursor-not-allowed"
              }`}
            >
              入力画面へ
            </button>
          </div>
        </div>
      )}

      {/* INPUT画面 */}
      {view === "INPUT" && (
        <div className="flex flex-1 flex-col p-3 md:p-4 max-w-5xl w-full mx-auto overflow-hidden">
          <header className="flex items-center justify-between bg-[#e1f3fb] px-4 py-3 rounded-t-xl border-b border-slate-200 shrink-0">
            <button
              onClick={() => { setView("HOME"); setSelectedUser(""); }}
              className="rounded-xl bg-white px-6 py-3 text-lg font-bold text-slate-700 shadow-sm border border-slate-200 active:scale-95 transition"
            >
              HOME
            </button>
            <div className="flex flex-col items-center">
              <div className="flex items-center gap-4 text-2xl font-black text-slate-800">
                <button onClick={() => changeDate(-1)} className="p-1 active:scale-125 transition"><ChevronLeft className="h-10 w-10" /></button>
                <span className="w-48 text-center text-xl md:text-2xl">{formattedDate}</span>
                <button onClick={() => changeDate(1)} className="p-1 active:scale-125 transition"><ChevronRight className="h-10 w-10" /></button>
              </div>
              <div className="mt-1 text-lg font-bold text-slate-700 underline decoration-sky-400 decoration-2 underline-offset-4">
                {selectedUser}
              </div>
            </div>
            <button
              onClick={() => {
                setView("HISTORY");
                fetchMonthlyHistory(selectedUser, currentDate);
              }}
              className="rounded-xl bg-white p-3 shadow-md border border-slate-200 text-orange-400 hover:bg-orange-50 active:scale-95 transition"
              title="履歴を見る"
            >
              <CalendarDays className="h-8 w-8" />
            </button>
          </header>

          <main className="flex flex-1 bg-white rounded-b-xl shadow-sm border-x border-b border-slate-200 p-4 gap-6 overflow-hidden">
            {/* 左側ツールバー */}
            <div className="flex w-36 flex-col items-center justify-center gap-5 border-r border-slate-100 pr-4 shrink-0">
              <button
                onClick={() => setActiveTool("ok")}
                className={`flex h-20 w-20 items-center justify-center rounded-full border-4 transition active:scale-90 ${
                  activeTool === "ok"
                    ? "border-blue-500 bg-blue-50 shadow-inner"
                    : "border-slate-200 bg-white opacity-40"
                }`}
              >
                <OkMark size="medium" />
              </button>
              <button
                onClick={() => setActiveTool("ng")}
                className={`flex h-20 w-20 items-center justify-center rounded-full border-4 transition active:scale-90 ${
                  activeTool === "ng"
                    ? "border-red-500 bg-red-50 shadow-inner"
                    : "border-slate-200 bg-white opacity-40"
                }`}
              >
                <NgMark size="medium" />
              </button>
              <button
                onClick={() => setActiveTool("clear")}
                className={`flex flex-col items-center justify-center gap-1 rounded-xl px-4 py-3 font-bold w-full transition active:scale-95 ${
                  activeTool === "clear" ? "bg-slate-200 text-slate-800" : "bg-slate-50 text-slate-400"
                }`}
              >
                <Eraser className="h-7 w-7" />
                <span className="text-sm">消しゴム</span>
              </button>
            </div>

            {/* 右側メインテーブル入力 */}
            <div className="flex-1 flex flex-col justify-between overflow-hidden">
              <div className="border border-slate-400 rounded-lg overflow-hidden">
                {/* ヘッダー行 */}
                <div className="grid grid-cols-4 border-b border-slate-400 bg-slate-100 text-center text-lg font-bold text-slate-700">
                  <div className="py-3 bg-slate-200/60 border-r border-slate-400"></div>
                  {MEAL_COLUMNS.map((col, i) => (
                    <div key={col} className={`py-3 text-xl ${i < 2 ? "border-r border-slate-400" : ""}`}>{col}</div>
                  ))}
                </div>
                
                {/* --- 判定行 --- */}
                <div className="grid grid-cols-4 border-b border-slate-400 items-center">
                  <div className="flex h-24 items-center justify-center bg-slate-50 font-bold border-r border-slate-400 text-slate-700">判定</div>
                  {tableData.map((data, idx) => (
                    <button
                      key={idx}
                      onClick={() => onJudgmentPress(idx)}
                      className={`flex h-24 items-center justify-center bg-white hover:bg-slate-50 transition ${idx < 2 ? "border-r border-slate-400" : ""}`}
                    >
                      {data.judgment === "ok" && <OkMark size="large" />}
                      {data.judgment === "ng" && <NgMark size="large" />}
                    </button>
                  ))}
                </div>

                {/* 主食カウント行 */}
                <div className="grid grid-cols-4 border-b border-slate-400 items-center">
                  <div className="flex h-24 items-center justify-center bg-slate-50 px-2 text-center text-base font-bold border-r border-slate-400 text-slate-700 leading-tight">
                    ライス/パン<br />＆スープ
                  </div>
                  {tableData.map((data, idx) => (
                    <button
                      key={idx}
                      onClick={() => onCountPress(idx)}
                      className={`flex h-24 items-center justify-center bg-white hover:bg-slate-50 text-4xl font-black text-slate-800 transition ${idx < 2 ? "border-r border-slate-400" : ""}`}
                    >
                      {data.count}
                    </button>
                  ))}
                </div>

                {/* メモ行 */}
                <div className="grid grid-cols-4 items-center">
                  <div className="flex h-24 items-center justify-center bg-slate-50 font-bold border-r border-slate-400 text-slate-700">メモ</div>
                  {tableData.map((data, idx) => (
                    <div key={idx} className={`h-24 p-1.5 bg-white ${idx < 2 ? "border-r border-slate-400" : ""}`}>
                      <textarea
                        value={data.memo}
                        onChange={(e) => {
                          const next = [...tableData];
                          next[idx] = { ...next[idx], memo: e.target.value };
                          setTableData(next);
                        }}
                        placeholder="メモを入力..."
                        className="h-full w-full resize-none p-2 text-base focus:outline-none bg-slate-50/50 rounded border border-slate-200"
                      />
                    </div>
                  ))}
                </div>
              </div>

              {/* 送信ボタン */}
              <div className="flex justify-center pt-4 shrink-0">
                <button
                  onClick={onSend}
                  disabled={isSending}
                  className="w-64 rounded-xl bg-sky-400 py-4 text-2xl font-bold text-white shadow-md active:scale-95 transition disabled:opacity-50"
                >
                  {isSending ? "送信中…" : "送信"}
                </button>
              </div>
            </div>
          </main>
        </div>
      )}

      {/* HISTORY画面 */}
      {view === "HISTORY" && (
        <div className="flex flex-1 flex-col p-3 md:p-4 max-w-5xl w-full mx-auto overflow-hidden">
          <header className="flex items-center justify-between bg-[#e1f3fb] px-4 py-3 rounded-t-xl border-b border-slate-200 shrink-0">
            <button
              onClick={() => setView("INPUT")}
              className="rounded-xl bg-white px-5 py-3 text-base md:text-lg font-bold text-slate-700 shadow-sm border border-slate-200 active:scale-95 transition shrink-0"
            >
              入力画面に戻る
            </button>
            <div className="flex-1 text-center font-black text-lg md:text-xl text-slate-800 truncate px-2">
              {selectedUser.replace("　", " ")} 履歴一覧
            </div>
            <button
              onClick={() => { setView("HOME"); setSelectedUser(""); }}
              className="flex items-center gap-1 rounded-xl bg-white px-5 py-3 text-base md:text-lg font-bold text-slate-700 shadow-sm border border-slate-200 active:scale-95 transition shrink-0"
            >
              <Home className="h-5 w-5 text-slate-500" />
              <span>HOME</span>
            </button>
          </header>

          <main className="flex-1 bg-white rounded-b-xl shadow-sm border-x border-b border-slate-200 p-3 overflow-y-auto">
            <div className="overflow-hidden border border-slate-400 rounded-lg">
              <table className="w-full border-collapse">
                <thead>
                  <tr className="divide-x divide-slate-400 border-b border-slate-400 bg-slate-100 font-bold text-slate-700">
                    <th className="py-3 w-28 text-center bg-slate-200/50 text-sm md:text-base">日付</th>
                    <th className="py-3 text-center text-sm md:text-base">朝</th>
                    <th className="py-3 text-center text-sm md:text-base">昼</th>
                    <th className="py-3 text-center text-sm md:text-base">夜</th>
                    <th className="py-3 text-center text-xs md:text-sm leading-tight">主食<br/>回数</th>
                    <th className="py-3 text-left px-4 text-sm md:text-base">メモ</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-300">
                  {monthlyHistory.map((day, idx) => (
                    <tr key={idx} className="divide-x divide-slate-300 text-center hover:bg-slate-50/80 transition h-14">
                      <td className="bg-slate-50 font-bold px-1 text-xs md:text-sm text-slate-600">{day.dateLabel}</td>
                      <td className="p-1">{day.morning === "ok" ? <OkMark size="small" /> : day.morning === "ng" ? <NgMark size="small" /> : ""}</td>
                      <td className="p-1">{day.lunch === "ok" ? <OkMark size="small" /> : day.lunch === "ng" ? <NgMark size="small" /> : ""}</td>
                      <td className="p-1">{day.dinner === "ok" ? <OkMark size="small" /> : day.dinner === "ng" ? <NgMark size="small" /> : ""}</td>
                      <td className="font-extrabold text-xl text-slate-800">{day.stapleCount}</td>
                      <td className="text-left px-4 text-xs md:text-sm whitespace-pre-wrap text-slate-600 max-w-xs truncate">{day.memo || "-"}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </main>
        </div>
      )}
    </div>
  );
}
