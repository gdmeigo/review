import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  Plus, Trash2, Pencil, ArrowLeft, Check, X, Play,
  Flame, Star, ChevronRight, BookOpen, Loader2, RotateCcw,
  Home as HomeIcon, Lock, Unlock, KeyRound, Smile, ImageOff,
  FileSpreadsheet, RefreshCw, Link as LinkIcon
} from "lucide-react";

/* ---------------------------------------------------------
   GDM Picture Review — a card-catalog styled flashcard drill
   for reviewing Richards' Graded Direct Method / English
   Through Pictures lessons. Picture-first, no translation —
   the review screen itself carries no Japanese instructions.

   Content (lessons & cards) is SHARED across everyone using
   this artifact. Each person's own review progress and XP
   stay PRIVATE to them. Editing requires the teacher
   passphrase. The teacher can either add cards by hand, or
   paste a link to a published Google Sheet / Excel CSV and
   sync — images are just URLs typed into a spreadsheet
   column, no upload needed.
--------------------------------------------------------- */

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,500;0,9..144,700;0,9..144,900;1,9..144,600&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');`;

const INTERVAL_DAYS = [0, 1, 2, 4, 7, 14]; // by mastery level 0-5
const DAY_MS = 24 * 60 * 60 * 1000;
const uid = () => Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
const todayStr = () => new Date().toISOString().slice(0, 10);
const emptyProgress = () => ({ level: 0, dueAt: 0, lastReview: 0 });
const DRIVE_IMAGE_SIZE = "w1000";

function getGoogleDriveFileId(url) {
  try {
    const parsed = new URL(url);
    const host = parsed.hostname.toLowerCase();
    if (host === "lh3.googleusercontent.com") {
      const directMatch = parsed.pathname.match(/^\/d\/([^/=]+)/);
      return directMatch ? directMatch[1] : "";
    }
    if (host !== "drive.google.com" && host !== "docs.google.com") return "";

    const fileMatch = parsed.pathname.match(/\/file\/d\/([^/]+)/);
    if (fileMatch) return fileMatch[1];

    const id = parsed.searchParams.get("id");
    return id || "";
  } catch {
    return "";
  }
}

function normalizeImageUrl(url) {
  const trimmed = (url || "").trim();
  if (!trimmed) return "";

  const driveFileId = getGoogleDriveFileId(trimmed);
  if (driveFileId) {
    return `https://lh3.googleusercontent.com/d/${encodeURIComponent(driveFileId)}=${DRIVE_IMAGE_SIZE}`;
  }

  return trimmed;
}

function normalizeSheetUrl(url) {
  const trimmed = (url || "").trim();
  const driveFileId = getGoogleDriveFileId(trimmed);
  if (driveFileId) {
    return `https://drive.usercontent.google.com/download?id=${encodeURIComponent(driveFileId)}&export=download`;
  }

  return trimmed;
}

const EMOJI_BANK = {
  "人": ["🧑", "👦", "👧", "👨", "👩", "👴", "👵", "🧑‍🏫", "🧑‍🎓", "👶"],
  "もの": ["📕", "✏️", "🖊️", "🪑", "🚪", "🪟", "🗝️", "🕰️", "🍎", "🍞", "🥛", "☕", "🎩", "👞", "🪆", "✉️"],
  "場所": ["🏠", "🏫", "🌳", "🛣️", "🌉", "🏞️", "🚉", "🏢"],
  "動作": ["🚶", "🏃", "🖐️", "👉", "🤲", "🪑", "📖", "✍️", "🍽️", "🚪", "🎁", "🤝"],
  "位置/前置詞": ["⬆️", "⬇️", "⬅️", "➡️", "🔼", "🔽", "🔁", "↔️", "🕳️", "📦"],
  "数": ["1️⃣", "2️⃣", "3️⃣", "4️⃣", "5️⃣", "6️⃣", "7️⃣", "8️⃣", "9️⃣", "🔟"],
};

/* ---------------- spreadsheet import helpers ---------------- */
const FIELD_ALIASES = {
  lesson: ["lesson", "レッスン", "レッスン名", "unit", "ユニット"],
  emoji: ["emoji", "絵文字"],
  en: ["english", "en", "英語", "英文", "word", "text", "単語"],
  image: ["image", "imageurl", "image_url", "画像", "画像url", "picture", "photo"],
  note: ["note", "メモ", "memo"],
};
const normKey = (k) => (k || "").toString().trim().toLowerCase().replace(/\s+/g, "");
function getField(rowObj, field) {
  const aliases = FIELD_ALIASES[field];
  for (const key of Object.keys(rowObj)) {
    if (aliases.some((a) => normKey(a) === normKey(key))) {
      const v = rowObj[key];
      return v == null ? "" : String(v).trim();
    }
  }
  return "";
}
function slugify(str) {
  const base = (str || "")
    .toString()
    .trim()
    .toLowerCase()
    .replace(/[^\p{L}\p{N}]+/gu, "-")
    .replace(/^-+|-+$/g, "");
  return base.slice(0, 40) || "x";
}
function buildContentFromRows(rows) {
  const order = [];
  const byTitle = {};
  rows.forEach((row) => {
    const en = getField(row, "en");
    if (!en) return;
    const title = getField(row, "lesson") || "未分類";
    const emoji = getField(row, "emoji");
    const image = normalizeImageUrl(getField(row, "image"));
    const note = getField(row, "note");
    if (!byTitle[title]) {
      byTitle[title] = { title, emoji: emoji || "📇", cards: [] };
      order.push(title);
    } else if (emoji && byTitle[title].emoji === "📇") {
      byTitle[title].emoji = emoji;
    }
    byTitle[title].cards.push({ en, image, emoji, note });
  });

  const usedLessonIds = new Set();
  const lessonsMap = {};
  const index = [];
  order.forEach((title) => {
    const l = byTitle[title];
    let lessonId = "l-" + slugify(title);
    let n = 2;
    while (usedLessonIds.has(lessonId)) lessonId = "l-" + slugify(title) + "-" + n++;
    usedLessonIds.add(lessonId);

    const usedCardIds = new Set();
    const cards = l.cards.map((c) => {
      let cardId = lessonId + "-" + slugify(c.en);
      let m = 2;
      while (usedCardIds.has(cardId)) cardId = lessonId + "-" + slugify(c.en) + "-" + m++;
      usedCardIds.add(cardId);
      return {
        id: cardId,
        en: c.en,
        note: c.note,
        visualType: c.image ? "photo" : "emoji",
        photoUrl: c.image || undefined,
        emoji: c.image ? "" : c.emoji || "❓",
      };
    });
    lessonsMap[lessonId] = { id: lessonId, title, emoji: l.emoji, cards };
    index.push({ id: lessonId, title, emoji: l.emoji, count: cards.length });
  });
  return { index, lessonsMap };
}
function isExcelBuffer(buffer) {
  const bytes = new Uint8Array(buffer.slice(0, 4));
  return bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
}

async function fetchAndParseSheet(url) {
  const fetchUrl = normalizeSheetUrl(url);
  const res = await fetch(fetchUrl, { mode: "cors" });
  if (!res.ok) throw new Error("http " + res.status);

  const lower = url.toLowerCase().split("?")[0];
  const contentType = (res.headers.get("content-type") || "").toLowerCase();
  const buf = await res.arrayBuffer();
  const looksLikeExcel =
    lower.endsWith(".xlsx") ||
    lower.endsWith(".xls") ||
    contentType.includes("spreadsheet") ||
    contentType.includes("excel") ||
    isExcelBuffer(buf);

  if (looksLikeExcel) {
    const wb = XLSX.read(buf, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: "" });
  }

  const text = new TextDecoder("utf-8").decode(buf).replace(/^\uFEFF/, "");
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  return parsed.data;
}

export default function App() {
  const [ready, setReady] = useState(false);
  const [lessons, setLessons] = useState({});
  const [index, setIndex] = useState([]);
  const [stats, setStats] = useState({ xp: 0, reviewDates: [] });
  const [progressByLesson, setProgressByLesson] = useState({});
  const [adminPinExists, setAdminPinExists] = useState(false);
  const [adminPinValue, setAdminPinValue] = useState(null);
  const [isEditor, setIsEditor] = useState(false);
  const [sheetUrl, setSheetUrl] = useState("");
  const [screen, setScreen] = useState({ name: "home" });
  const [pinModal, setPinModal] = useState(null);
  const [toast, setToast] = useState(null);

  const showToast = useCallback((msg) => {
    setToast(msg);
    setTimeout(() => setToast(null), 2600);
  }, []);

  const safeGet = async (key, shared) => {
    try {
      const r = await window.storage.get(key, shared);
      return r ? r.value : null;
    } catch {
      return null;
    }
  };
  const safeSet = async (key, value, shared) => {
    try {
      await window.storage.set(key, value, shared);
      return true;
    } catch {
      return false;
    }
  };
  const safeDelete = async (key, shared) => {
    try {
      await window.storage.delete(key, shared);
    } catch {}
  };

  useEffect(() => {
    (async () => {
      const idxRaw = await safeGet("lesson-index", true);
      const idx = idxRaw ? JSON.parse(idxRaw) : [];
      const statsRaw = await safeGet("stats", false);
      const st = statsRaw ? JSON.parse(statsRaw) : { xp: 0, reviewDates: [] };
      const editorRaw = await safeGet("is-editor", false);
      const pinRaw = await safeGet("admin-pin", true);
      const sheetRaw = await safeGet("sheet-url", true);

      const loaded = {};
      for (const meta of idx) {
        const raw = await safeGet("lesson:" + meta.id, true);
        if (raw) loaded[meta.id] = JSON.parse(raw);
      }
      setIndex(idx);
      setLessons(loaded);
      setStats(st);
      setIsEditor(editorRaw === "true");
      setAdminPinExists(!!pinRaw);
      setAdminPinValue(pinRaw);
      setSheetUrl(sheetRaw || "");
      setReady(true);
    })();
  }, []);

  const persistStats = useCallback(async (s) => {
    setStats(s);
    await safeSet("stats", JSON.stringify(s), false);
  }, []);

  const loadProgress = useCallback(
    async (lessonId) => {
      if (progressByLesson[lessonId]) return progressByLesson[lessonId];
      const raw = await safeGet("progress:" + lessonId, false);
      const prog = raw ? JSON.parse(raw) : {};
      setProgressByLesson((prev) => ({ ...prev, [lessonId]: prog }));
      return prog;
    },
    [progressByLesson]
  );

  const persistProgress = useCallback(async (lessonId, prog) => {
    setProgressByLesson((prev) => ({ ...prev, [lessonId]: prog }));
    await safeSet("progress:" + lessonId, JSON.stringify(prog), false);
  }, []);

  // ---- editor / PIN flow ----
  const requestEditorAccess = () => setPinModal(adminPinExists ? "enter" : "setup");

  const submitSetupPin = async (pin) => {
    await safeSet("admin-pin", pin, true);
    setAdminPinExists(true);
    setAdminPinValue(pin);
    await safeSet("is-editor", "true", false);
    setIsEditor(true);
    setPinModal(null);
    showToast("合言葉を作成し、編集モードにしました");
  };
  const submitEnterPin = async (pin) => {
    const current = adminPinValue ?? (await safeGet("admin-pin", true));
    if (pin === current) {
      await safeSet("is-editor", "true", false);
      setIsEditor(true);
      setPinModal(null);
      showToast("編集モードにしました");
    } else {
      showToast("合言葉が違います");
    }
  };
  const exitEditorMode = async () => {
    await safeSet("is-editor", "false", false);
    setIsEditor(false);
    showToast("編集モードを終了しました");
  };

  // ---- manual lesson / card mutations ----
  const createLesson = async (title, emoji) => {
    const id = uid();
    const lesson = { id, title, emoji, cards: [] };
    const newIdx = [...index, { id, title, emoji, count: 0 }];
    setIndex(newIdx);
    setLessons((prev) => ({ ...prev, [id]: lesson }));
    await safeSet("lesson-index", JSON.stringify(newIdx), true);
    await safeSet("lesson:" + id, JSON.stringify(lesson), true);
    return id;
  };
  const updateLessonMeta = async (id, title, emoji) => {
    const lesson = { ...lessons[id], title, emoji };
    setLessons((prev) => ({ ...prev, [id]: lesson }));
    await safeSet("lesson:" + id, JSON.stringify(lesson), true);
    const newIdx = index.map((m) => (m.id === id ? { ...m, title, emoji } : m));
    setIndex(newIdx);
    await safeSet("lesson-index", JSON.stringify(newIdx), true);
  };
  const deleteLesson = async (id) => {
    await safeDelete("lesson:" + id, true);
    await safeDelete("progress:" + id, false);
    const newIdx = index.filter((m) => m.id !== id);
    setIndex(newIdx);
    await safeSet("lesson-index", JSON.stringify(newIdx), true);
    setLessons((prev) => {
      const n = { ...prev };
      delete n[id];
      return n;
    });
  };
  const upsertCard = async (lessonId, card) => {
    const lesson = lessons[lessonId];
    const exists = lesson.cards.some((c) => c.id === card.id);
    const cards = exists ? lesson.cards.map((c) => (c.id === card.id ? card : c)) : [...lesson.cards, card];
    const updated = { ...lesson, cards };
    setLessons((prev) => ({ ...prev, [lessonId]: updated }));
    await safeSet("lesson:" + lessonId, JSON.stringify(updated), true);
    const newIdx = index.map((m) => (m.id === lessonId ? { ...m, count: cards.length } : m));
    setIndex(newIdx);
    await safeSet("lesson-index", JSON.stringify(newIdx), true);
  };
  const deleteCard = async (lessonId, cardId) => {
    const lesson = lessons[lessonId];
    const cards = lesson.cards.filter((c) => c.id !== cardId);
    const updated = { ...lesson, cards };
    setLessons((prev) => ({ ...prev, [lessonId]: updated }));
    await safeSet("lesson:" + lessonId, JSON.stringify(updated), true);
    const newIdx = index.map((m) => (m.id === lessonId ? { ...m, count: cards.length } : m));
    setIndex(newIdx);
    await safeSet("lesson-index", JSON.stringify(newIdx), true);
  };

  // ---- spreadsheet sync ----
  const saveSheetUrl = async (url) => {
    setSheetUrl(url);
    await safeSet("sheet-url", url, true);
  };
  const syncFromSheet = async (url) => {
    const rows = await fetchAndParseSheet(url);
    const { index: newIndex, lessonsMap: newLessons } = buildContentFromRows(rows);
    for (const oldMeta of index) {
      if (!newLessons[oldMeta.id]) await safeDelete("lesson:" + oldMeta.id, true);
    }
    for (const l of Object.values(newLessons)) {
      await safeSet("lesson:" + l.id, JSON.stringify(l), true);
    }
    await safeSet("lesson-index", JSON.stringify(newIndex), true);
    setIndex(newIndex);
    setLessons(newLessons);
    return newIndex;
  };

  const allCardsFlat = useMemo(
    () => Object.values(lessons).flatMap((l) => l.cards.map((c) => ({ ...c, lessonId: l.id }))),
    [lessons]
  );

  const streak = useMemo(() => {
    const dates = new Set(stats.reviewDates || []);
    let n = 0;
    let cursor = new Date();
    if (!dates.has(todayStr())) cursor = new Date(Date.now() - DAY_MS);
    while (dates.has(cursor.toISOString().slice(0, 10))) {
      n++;
      cursor = new Date(cursor.getTime() - DAY_MS);
    }
    return n;
  }, [stats.reviewDates]);

  return (
    <div style={{ fontFamily: "'Inter', sans-serif" }} className="min-h-screen w-full">
      <style>{`
        ${FONT_IMPORT}
        .font-display { font-family: 'Fraunces', serif; }
        .font-mono { font-family: 'IBM Plex Mono', monospace; }
        .cabinet-bg {
          background-color: #2E2117;
          background-image:
            repeating-linear-gradient(90deg, rgba(0,0,0,0.12) 0px, rgba(0,0,0,0.12) 2px, transparent 2px, transparent 90px),
            linear-gradient(180deg, #3B2A1C 0%, #2A1D13 100%);
        }
        .card-paper {
          background: #F3ECDA;
          background-image: radial-gradient(rgba(0,0,0,0.035) 1px, transparent 1px);
          background-size: 6px 6px;
        }
        .punch-hole { width: 14px; height: 14px; border-radius: 50%; background: #2E2117; box-shadow: inset 0 2px 3px rgba(0,0,0,0.6); }
        .drawer-front { background: linear-gradient(180deg, #6B4A31 0%, #57381F 100%); border: 1px solid #402A17; }
        .brass { background: linear-gradient(180deg, #E3B355 0%, #B9853A 100%); border: 1px solid #8C6425; }
        @keyframes stampIn { 0% { transform: scale(2.2) rotate(-12deg); opacity: 0; } 60% { transform: scale(0.95) rotate(-12deg); opacity: 1; } 100% { transform: scale(1) rotate(-12deg); opacity: 1; } }
        .stamp { animation: stampIn 0.35s ease-out; }
        @keyframes slideUp { from { transform: translateY(14px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
        .slide-up { animation: slideUp 0.28s ease-out; }
        @media (prefers-reduced-motion: reduce) { .stamp, .slide-up { animation: none; } }
        button:focus-visible, input:focus-visible, textarea:focus-visible { outline: 3px solid #E3B355; outline-offset: 2px; }
      `}</style>

      {!ready ? (
        <div className="cabinet-bg min-h-screen flex items-center justify-center">
          <Loader2 className="animate-spin text-[#E3B355]" size={32} />
        </div>
      ) : (
        <div className="cabinet-bg min-h-screen">
          <TopBar screen={screen} setScreen={setScreen} xp={stats.xp} streak={streak} isEditor={isEditor} onRequestEditor={requestEditorAccess} onExitEditor={exitEditorMode} />
          <div className="max-w-2xl mx-auto px-4 pb-16">
            {screen.name === "home" && (
              <Home
                index={index}
                isEditor={isEditor}
                sheetUrl={sheetUrl}
                onSaveSheetUrl={saveSheetUrl}
                onSyncSheet={syncFromSheet}
                onOpen={(id) => setScreen({ name: "lesson", id })}
                onCreate={async (title, emoji) => {
                  const id = await createLesson(title, emoji);
                  setScreen({ name: "lesson", id });
                }}
              />
            )}
            {screen.name === "lesson" && lessons[screen.id] && (
              <LessonDetail
                lesson={lessons[screen.id]}
                isEditor={isEditor}
                progress={progressByLesson[screen.id]}
                onEnsureProgress={() => loadProgress(screen.id)}
                onBack={() => setScreen({ name: "home" })}
                onUpdateMeta={(t, e) => updateLessonMeta(screen.id, t, e)}
                onDeleteLesson={async () => {
                  await deleteLesson(screen.id);
                  setScreen({ name: "home" });
                }}
                onUpsertCard={(card) => upsertCard(screen.id, card)}
                onDeleteCard={(cid) => deleteCard(screen.id, cid)}
                onStartReview={() => setScreen({ name: "review", id: screen.id })}
              />
            )}
            {screen.name === "review" && lessons[screen.id] && (
              <ReviewSession
                lesson={lessons[screen.id]}
                allCards={allCardsFlat}
                initialProgress={progressByLesson[screen.id] || {}}
                onFinish={async (result) => {
                  const prevProgress = progressByLesson[screen.id] || {};
                  const nextProgress = { ...prevProgress };
                  for (const u of result.updates) nextProgress[u.id] = { level: u.level, dueAt: u.dueAt, lastReview: u.lastReview };
                  await persistProgress(screen.id, nextProgress);
                  const dates = new Set(stats.reviewDates || []);
                  dates.add(todayStr());
                  await persistStats({ xp: (stats.xp || 0) + result.xpEarned, reviewDates: Array.from(dates) });
                  setScreen({ name: "summary", id: screen.id, result });
                }}
                onExit={() => setScreen({ name: "lesson", id: screen.id })}
              />
            )}
            {screen.name === "summary" && (
              <Summary result={screen.result} onReviewAgain={() => setScreen({ name: "review", id: screen.id })} onBackToLesson={() => setScreen({ name: "lesson", id: screen.id })} onHome={() => setScreen({ name: "home" })} />
            )}
          </div>
        </div>
      )}

      {pinModal && <PinModal mode={pinModal} onCancel={() => setPinModal(null)} onSubmitSetup={submitSetupPin} onSubmitEnter={submitEnterPin} />}

      {toast && <div className="fixed bottom-5 left-1/2 -translate-x-1/2 bg-[#1F2A3C] text-[#F3ECDA] px-4 py-2 rounded-md text-sm font-medium shadow-lg z-50 slide-up">{toast}</div>}
    </div>
  );
}

/* ---------------- TopBar ---------------- */
function TopBar({ screen, setScreen, xp, streak, isEditor, onRequestEditor, onExitEditor }) {
  return (
    <div className="sticky top-0 z-30 backdrop-blur-sm bg-[#2E2117]/90 border-b border-[#4a3520]">
      <div className="max-w-2xl mx-auto px-4 py-3 flex items-center justify-between gap-2">
        <button onClick={() => setScreen({ name: "home" })} className="flex items-center gap-2 text-[#F3ECDA] min-w-0">
          <span className="font-display text-lg font-bold tracking-tight truncate">GDM 復習カード</span>
        </button>
        <div className="flex items-center gap-3 font-mono text-sm shrink-0">
          <div className="flex items-center gap-1 text-[#E3B355]"><Flame size={16} /> {streak}</div>
          <div className="flex items-center gap-1 text-[#D9C79A]"><Star size={16} /> {xp}</div>
          {isEditor ? (
            <button onClick={onExitEditor} className="flex items-center gap-1 text-[10px] bg-[#2F6B4F] text-[#F3ECDA] px-2 py-1 rounded" title="編集モードを終了">
              <Unlock size={12} /> 編集中
            </button>
          ) : (
            <button onClick={onRequestEditor} className="flex items-center gap-1 text-[10px] bg-[#5a4a30] text-[#D9C79A] px-2 py-1 rounded hover:text-[#F3ECDA]" title="先生モードに入る">
              <Lock size={12} /> 先生
            </button>
          )}
          {screen.name !== "home" && (
            <button
              onClick={() => setScreen(screen.name === "summary" ? { name: "home" } : screen.name === "review" ? { name: "lesson", id: screen.id } : { name: "home" })}
              className="text-[#D9C79A] hover:text-[#F3ECDA]"
              aria-label="戻る"
            >
              <ArrowLeft size={18} />
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

/* ---------------- PIN Modal ---------------- */
function PinModal({ mode, onCancel, onSubmitSetup, onSubmitEnter }) {
  const [pin, setPin] = useState("");
  const [pin2, setPin2] = useState("");
  const [error, setError] = useState("");
  const handleSubmit = () => {
    if (mode === "setup") {
      if (pin.length < 4) return setError("4文字以上にしてください");
      if (pin !== pin2) return setError("合言葉が一致しません");
      onSubmitSetup(pin);
    } else {
      if (!pin) return;
      onSubmitEnter(pin);
    }
  };
  return (
    <div className="fixed inset-0 bg-black/60 flex items-center justify-center z-50 p-4">
      <div className="card-paper rounded-md p-5 border border-[#c9bb9a] w-full max-w-sm slide-up">
        <div className="flex items-center gap-2 mb-3 text-[#3B2A1C]">
          <KeyRound size={18} />
          <h3 className="font-display text-lg font-bold">{mode === "setup" ? "先生用の合言葉を作成" : "先生モードに入る"}</h3>
        </div>
        {mode === "setup" && (
          <p className="text-xs text-[#6b5d44] mb-3">
            まだ合言葉が設定されていません。ここで決めた合言葉を知っている人だけがレッスン内容を編集できます。（簡易的な仕組みのため、厳密なセキュリティではありません）
          </p>
        )}
        <input type="password" value={pin} onChange={(e) => { setPin(e.target.value); setError(""); }} placeholder="合言葉" autoFocus className="w-full rounded border border-[#c9bb9a] bg-white/70 px-3 py-2 mb-2 font-mono" />
        {mode === "setup" && (
          <input type="password" value={pin2} onChange={(e) => { setPin2(e.target.value); setError(""); }} placeholder="もう一度入力" className="w-full rounded border border-[#c9bb9a] bg-white/70 px-3 py-2 mb-2 font-mono" />
        )}
        {error && <p className="text-xs text-[#a3402f] mb-2">{error}</p>}
        <div className="flex gap-2 justify-end mt-2">
          <button onClick={onCancel} className="px-3 py-1.5 text-sm text-[#6b5d44]">キャンセル</button>
          <button onClick={handleSubmit} className="px-4 py-1.5 text-sm rounded bg-[#1F2A3C] text-[#F3ECDA]">{mode === "setup" ? "作成して編集モードへ" : "入る"}</button>
        </div>
      </div>
    </div>
  );
}

/* ---------------- Home ---------------- */
function Home({ index, isEditor, sheetUrl, onSaveSheetUrl, onSyncSheet, onOpen, onCreate }) {
  const [showAdd, setShowAdd] = useState(false);
  return (
    <div className="pt-6">
      <p className="font-display italic text-[#D9C79A] mb-2 text-[15px]">Look at the picture. Say it in English.</p>
      <p className="text-[11px] text-[#8a7a5c] mb-5 font-mono">レッスン内容はクラス全員で共有されます。復習の記録(習熟度・XP)は自分だけに保存されます。</p>

      {index.length === 0 && !showAdd && (
        <div className="card-paper rounded-md p-6 text-center mb-4 border border-[#c9bb9a]">
          <p className="font-display text-lg text-[#3B2A1C] mb-1">まだレッスンがありません</p>
          <p className="text-sm text-[#6b5d44]">{isEditor ? "下のスプレッドシート連携、または手動追加で作りましょう" : "先生がレッスンを追加するまでお待ちください"}</p>
        </div>
      )}

      <div className="space-y-3">
        {index.map((meta) => (
          <button key={meta.id} onClick={() => onOpen(meta.id)} className="drawer-front w-full rounded-md p-4 flex items-center gap-4 text-left shadow-md hover:brightness-110 transition">
            <div className="brass rounded w-12 h-12 flex items-center justify-center text-2xl shrink-0">{meta.emoji || "📇"}</div>
            <div className="flex-1 min-w-0">
              <div className="font-display text-[#F3ECDA] text-lg font-semibold truncate">{meta.title}</div>
              <div className="font-mono text-xs text-[#D9C79A]">{meta.count || 0} 枚のカード</div>
            </div>
            <ChevronRight className="text-[#D9C79A]" size={20} />
          </button>
        ))}
      </div>

      {isEditor && (
        <>
          {showAdd ? (
            <AddLessonForm onCancel={() => setShowAdd(false)} onCreate={(t, e) => { onCreate(t, e); setShowAdd(false); }} />
          ) : (
            <button onClick={() => setShowAdd(true)} className="mt-4 w-full rounded-md p-4 flex items-center justify-center gap-2 border-2 border-dashed border-[#6b5d44] text-[#D9C79A] hover:text-[#F3ECDA] hover:border-[#E3B355] transition">
              <Plus size={18} /> レッスンを手動で追加
            </button>
          )}
          <SheetSyncPanel sheetUrl={sheetUrl} onSaveSheetUrl={onSaveSheetUrl} onSyncSheet={onSyncSheet} />
        </>
      )}
    </div>
  );
}

function AddLessonForm({ onCancel, onCreate }) {
  const [title, setTitle] = useState("");
  const [emoji, setEmoji] = useState("📇");
  return (
    <div className="card-paper rounded-md p-4 mt-4 border border-[#c9bb9a] slide-up">
      <div className="flex gap-3 mb-3">
        <input value={emoji} onChange={(e) => setEmoji(e.target.value.slice(0, 2))} className="w-14 h-14 text-2xl text-center rounded border border-[#c9bb9a] bg-white/60" aria-label="レッスンのアイコン絵文字" />
        <input autoFocus value={title} onChange={(e) => setTitle(e.target.value)} placeholder="レッスン名（例: Lesson 1 — This is a book）" className="flex-1 rounded border border-[#c9bb9a] bg-white/60 px-3 text-[#3B2A1C] font-display" />
      </div>
      <div className="flex gap-2 justify-end">
        <button onClick={onCancel} className="px-3 py-1.5 text-sm text-[#6b5d44]">キャンセル</button>
        <button disabled={!title.trim()} onClick={() => onCreate(title.trim(), emoji || "📇")} className="px-4 py-1.5 text-sm rounded bg-[#1F2A3C] text-[#F3ECDA] disabled:opacity-40">作成</button>
      </div>
    </div>
  );
}

/* ---------------- Sheet Sync Panel (teacher only) ---------------- */
function SheetSyncPanel({ sheetUrl, onSaveSheetUrl, onSyncSheet }) {
  const [url, setUrl] = useState(sheetUrl || "");
  const [status, setStatus] = useState(null);
  const [confirming, setConfirming] = useState(false);
  const [showHelp, setShowHelp] = useState(false);

  const doSync = async () => {
    setConfirming(false);
    if (!url.trim()) return;
    setStatus({ type: "loading", message: "読み込み中…" });
    try {
      await onSaveSheetUrl(url.trim());
      const newIndex = await onSyncSheet(url.trim());
      const totalCards = newIndex.reduce((s, m) => s + (m.count || 0), 0);
      setStatus({ type: "success", message: `${newIndex.length}レッスン・${totalCards}枚のカードを読み込みました` });
    } catch (e) {
      setStatus({ type: "error", message: "読み込みに失敗しました。URLが「ウェブに公開」されたCSVリンクか確認してください。" });
    }
  };

  return (
    <div className="card-paper rounded-md p-4 mt-4 border border-[#c9bb9a]">
      <div className="flex items-center gap-2 mb-2 text-[#3B2A1C]">
        <FileSpreadsheet size={18} />
        <h3 className="font-display font-bold">スプレッドシートから読み込み</h3>
      </div>
      <p className="text-xs text-[#6b5d44] mb-3">
        Googleスプレッドシート（またはExcel）を「ウェブに公開 → CSV」の形にして、そのリンクを貼ってください。列は
        <span className="font-mono"> lesson / english / image / emoji / note </span>
        （日本語見出しでも可：レッスン / 英語 / 画像 / 絵文字 / メモ）。画像は列にURLを貼るだけでOKです。
      </p>
      <button onClick={() => setShowHelp((s) => !s)} className="text-xs text-[#8C6425] underline mb-3">
        {showHelp ? "手順を隠す" : "公開リンクの作り方を見る"}
      </button>
      {showHelp && (
        <ol className="text-xs text-[#6b5d44] list-decimal list-inside mb-3 space-y-1">
          <li>Googleスプレッドシートを開く（Excelの場合は先にGoogleスプレッドシートにインポート）</li>
          <li>「ファイル」→「共有」→「ウェブに公開」</li>
          <li>公開する範囲でシートを選び、形式を「カンマ区切りの値(.csv)」にして公開</li>
          <li>表示されたURLをコピーして下に貼り付け</li>
        </ol>
      )}
      <div className="flex gap-2 mb-2">
        <LinkIcon size={16} className="text-[#6b5d44] shrink-0 mt-2.5" />
        <input
          value={url}
          onChange={(e) => setUrl(e.target.value)}
          placeholder="https://docs.google.com/spreadsheets/d/e/.../pub?output=csv"
          className="flex-1 rounded border border-[#c9bb9a] bg-white/60 px-3 py-2 text-sm font-mono"
        />
      </div>
      {!confirming ? (
        <button
          disabled={!url.trim() || status?.type === "loading"}
          onClick={() => setConfirming(true)}
          className="w-full rounded-md py-2 flex items-center justify-center gap-2 bg-[#1F2A3C] text-[#F3ECDA] font-display font-semibold disabled:opacity-40"
        >
          {status?.type === "loading" ? <Loader2 className="animate-spin" size={16} /> : <RefreshCw size={16} />}
          同期する
        </button>
      ) : (
        <div className="text-sm bg-[#f6e3dd] border border-[#e0b7ab] rounded p-2">
          <p className="text-[#7c2c1f] mb-2">同期すると現在の内容がスプレッドシートの内容に置き換わります（手動で追加した分も含む）。よろしいですか？</p>
          <div className="flex gap-2 justify-end">
            <button onClick={() => setConfirming(false)} className="text-[#6b5d44] px-2">やめる</button>
            <button onClick={doSync} className="text-[#7c2c1f] font-semibold px-2">同期する</button>
          </div>
        </div>
      )}
      {status && status.type !== "loading" && (
        <p className={`text-xs mt-2 ${status.type === "success" ? "text-[#2F6B4F]" : "text-[#a3402f]"}`}>{status.message}</p>
      )}
    </div>
  );
}

/* ---------------- Card Visual ---------------- */
function CardVisual({ card, className, iconSize = 18 }) {
  const [failedUrl, setFailedUrl] = useState("");

  if (card.visualType === "photo") {
    const photoUrl = normalizeImageUrl(card.photoUrl);
    if (!photoUrl || failedUrl === photoUrl) return <ImageOff className={className} size={iconSize} />;
    return (
      <img
        src={photoUrl}
        alt={card.en || ""}
        className={className}
        onError={() => setFailedUrl(photoUrl)}
      />
    );
  }
  return <span className={className}>{card.emoji || "🖼️"}</span>;
}

/* ---------------- Lesson Detail ---------------- */
function LessonDetail({ lesson, isEditor, progress, onEnsureProgress, onBack, onUpdateMeta, onDeleteLesson, onUpsertCard, onDeleteCard, onStartReview }) {
  const [editingMeta, setEditingMeta] = useState(false);
  const [title, setTitle] = useState(lesson.title);
  const [emoji, setEmoji] = useState(lesson.emoji);
  const [showCardForm, setShowCardForm] = useState(false);
  const [editingCard, setEditingCard] = useState(null);
  const [confirmDelete, setConfirmDelete] = useState(false);

  useEffect(() => {
    onEnsureProgress();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lesson.id]);

  const prog = progress || {};

  return (
    <div className="pt-6">
      <button onClick={onBack} className="flex items-center gap-1 text-[#D9C79A] text-sm mb-4 hover:text-[#F3ECDA]">
        <ArrowLeft size={16} /> 引き出し一覧へ
      </button>

      <div className="card-paper rounded-md p-4 border border-[#c9bb9a] mb-4">
        {editingMeta ? (
          <div className="flex gap-3">
            <input value={emoji} onChange={(e) => setEmoji(e.target.value.slice(0, 2))} className="w-12 h-12 text-xl text-center rounded border border-[#c9bb9a]" />
            <input value={title} onChange={(e) => setTitle(e.target.value)} className="flex-1 rounded border border-[#c9bb9a] px-3 font-display" />
            <button onClick={() => { onUpdateMeta(title.trim() || lesson.title, emoji || lesson.emoji); setEditingMeta(false); }} className="px-3 rounded bg-[#1F2A3C] text-[#F3ECDA] text-sm">保存</button>
          </div>
        ) : (
          <div className="flex items-center gap-3">
            <div className="brass rounded w-11 h-11 flex items-center justify-center text-xl shrink-0">{lesson.emoji}</div>
            <div className="flex-1">
              <div className="font-display text-xl font-bold text-[#3B2A1C]">{lesson.title}</div>
              <div className="font-mono text-xs text-[#6b5d44]">{lesson.cards.length} 枚</div>
            </div>
            {isEditor && (
              <>
                <button onClick={() => setEditingMeta(true)} className="text-[#6b5d44] hover:text-[#3B2A1C]" aria-label="レッスン名を編集"><Pencil size={16} /></button>
                <button onClick={() => setConfirmDelete(true)} className="text-[#a3402f] hover:text-[#7c2c1f]" aria-label="レッスンを削除"><Trash2 size={16} /></button>
              </>
            )}
          </div>
        )}
        {confirmDelete && (
          <div className="mt-3 text-sm bg-[#f6e3dd] border border-[#e0b7ab] rounded p-2 flex items-center justify-between">
            <span className="text-[#7c2c1f]">このレッスンを削除しますか？カードもすべて消えます。</span>
            <div className="flex gap-2 shrink-0 ml-2">
              <button onClick={() => setConfirmDelete(false)} className="text-[#6b5d44]">やめる</button>
              <button onClick={onDeleteLesson} className="text-[#7c2c1f] font-semibold">削除</button>
            </div>
          </div>
        )}
      </div>

      <button onClick={onStartReview} disabled={lesson.cards.length < 2} className="w-full mb-5 rounded-md py-3 flex items-center justify-center gap-2 bg-[#2F6B4F] text-[#F3ECDA] font-display font-bold text-lg shadow-md disabled:opacity-40 disabled:cursor-not-allowed hover:brightness-110 transition">
        <Play size={20} /> 復習をはじめる
      </button>
      {lesson.cards.length < 2 && <p className="text-xs text-[#D9C79A] -mt-3 mb-4">復習を始めるにはカードが2枚以上必要です</p>}

      <div className="space-y-2 mb-4">
        {lesson.cards.map((c) => (
          <div key={c.id} className="card-paper rounded-md p-3 border border-[#c9bb9a] flex items-center gap-3">
            <div className="punch-hole ml-1 hidden sm:block" />
            <div className="w-10 h-10 shrink-0 flex items-center justify-center text-2xl rounded overflow-hidden bg-white/40">
              <CardVisual card={c} className="w-10 h-10 object-cover rounded text-2xl flex items-center justify-center" iconSize={18} />
            </div>
            <div className="flex-1 min-w-0">
              <div className="font-display text-[#3B2A1C] font-semibold truncate">{c.en}</div>
              {c.note && <div className="text-xs text-[#6b5d44] truncate">{c.note}</div>}
            </div>
            <MasteryDots level={(prog[c.id] || emptyProgress()).level} />
            {isEditor && (
              <>
                <button onClick={() => setEditingCard(c)} className="text-[#6b5d44] hover:text-[#3B2A1C]" aria-label="編集"><Pencil size={15} /></button>
                <button onClick={() => onDeleteCard(c.id)} className="text-[#a3402f] hover:text-[#7c2c1f]" aria-label="削除"><Trash2 size={15} /></button>
              </>
            )}
          </div>
        ))}
      </div>

      {isEditor &&
        (showCardForm || editingCard ? (
          <CardForm
            initial={editingCard}
            onCancel={() => { setShowCardForm(false); setEditingCard(null); }}
            onSave={(card) => { onUpsertCard(card); setShowCardForm(false); setEditingCard(null); }}
          />
        ) : (
          <button onClick={() => setShowCardForm(true)} className="w-full rounded-md p-3 flex items-center justify-center gap-2 border-2 border-dashed border-[#6b5d44] text-[#D9C79A] hover:text-[#F3ECDA] hover:border-[#E3B355] transition">
            <Plus size={16} /> カードを手動で追加
          </button>
        ))}
    </div>
  );
}

function MasteryDots({ level }) {
  return (
    <div className="hidden sm:flex gap-0.5 mr-1" aria-label={`習熟度 ${level}/5`}>
      {[0, 1, 2, 3, 4].map((i) => <div key={i} className={`w-1.5 h-1.5 rounded-full ${i < level ? "bg-[#2F6B4F]" : "bg-[#c9bb9a]"}`} />)}
    </div>
  );
}

/* ---------------- Card Form (manual add, image = URL only) ---------------- */
function CardForm({ initial, onCancel, onSave }) {
  const [visualType, setVisualType] = useState(initial?.visualType || "emoji");
  const [emoji, setEmoji] = useState(initial?.emoji || "");
  const [photoUrl, setPhotoUrl] = useState(initial?.photoUrl || "");
  const [en, setEn] = useState(initial?.en || "");
  const [note, setNote] = useState(initial?.note || "");
  const [tab, setTab] = useState(Object.keys(EMOJI_BANK)[0]);
  const normalizedPhotoUrl = normalizeImageUrl(photoUrl);

  const canSave = en.trim() && ((visualType === "emoji" && emoji.trim()) || (visualType === "photo" && photoUrl.trim()));

  const handleSave = () => {
    onSave({
      id: initial?.id || uid(),
      en: en.trim(),
      note: note.trim(),
      visualType,
      emoji: visualType === "emoji" ? emoji.trim() : "",
      photoUrl: visualType === "photo" ? normalizedPhotoUrl : undefined,
    });
  };

  return (
    <div className="card-paper rounded-md p-4 border border-[#c9bb9a] slide-up">
      <div className="flex gap-1.5 mb-3">
        <button onClick={() => setVisualType("emoji")} className={`flex-1 text-sm px-3 py-2 rounded flex items-center justify-center gap-1.5 font-medium ${visualType === "emoji" ? "bg-[#1F2A3C] text-[#F3ECDA]" : "bg-white/50 text-[#6b5d44]"}`}>
          <Smile size={15} /> 絵文字
        </button>
        <button onClick={() => setVisualType("photo")} className={`flex-1 text-sm px-3 py-2 rounded flex items-center justify-center gap-1.5 font-medium ${visualType === "photo" ? "bg-[#1F2A3C] text-[#F3ECDA]" : "bg-white/50 text-[#6b5d44]"}`}>
          <LinkIcon size={15} /> 画像URL
        </button>
      </div>

      {visualType === "emoji" ? (
        <>
          <div className="flex gap-3 mb-3">
            <input value={emoji} onChange={(e) => setEmoji(e.target.value.slice(0, 4))} placeholder="🖼️" className="w-16 h-16 text-3xl text-center rounded border border-[#c9bb9a] bg-white/60" aria-label="絵（絵文字）" />
            <EnNoteInputs en={en} setEn={setEn} note={note} setNote={setNote} />
          </div>
          <div className="mb-3">
            <div className="flex flex-wrap gap-1 mb-2">
              {Object.keys(EMOJI_BANK).map((k) => (
                <button key={k} onClick={() => setTab(k)} className={`text-xs px-2 py-1 rounded font-mono ${tab === k ? "bg-[#1F2A3C] text-[#F3ECDA]" : "bg-white/50 text-[#6b5d44]"}`}>{k}</button>
              ))}
            </div>
            <div className="flex flex-wrap gap-1.5">
              {EMOJI_BANK[tab].map((em) => (
                <button key={em} onClick={() => setEmoji(em)} className="text-xl w-9 h-9 flex items-center justify-center rounded hover:bg-white/60 bg-white/30">{em}</button>
              ))}
            </div>
          </div>
        </>
      ) : (
        <>
          <div className="flex gap-3 mb-3">
            <div className="w-16 h-16 rounded border border-[#c9bb9a] bg-white/60 flex items-center justify-center overflow-hidden shrink-0">
              {normalizedPhotoUrl ? (
                <img src={normalizedPhotoUrl} alt="" className="w-16 h-16 object-cover" onError={(e) => (e.currentTarget.style.display = "none")} />
              ) : (
                <ImageOff className="text-[#8a7a5c]" size={20} />
              )}
            </div>
            <EnNoteInputs en={en} setEn={setEn} note={note} setNote={setNote} />
          </div>
          <input value={photoUrl} onChange={(e) => setPhotoUrl(e.target.value)} placeholder="https://example.com/apple.jpg" className="w-full rounded border border-[#c9bb9a] bg-white/60 px-3 py-2 text-sm mb-1" />
          <p className="text-[10px] text-[#8a7a5c] mb-2">画像はどこかにアップロード済みのものへのリンクを貼ってください。多くの場合はスプレッドシート連携でまとめて登録する方が簡単です。</p>
        </>
      )}

      <div className="flex gap-2 justify-end mt-2">
        <button onClick={onCancel} className="px-3 py-1.5 text-sm text-[#6b5d44]">キャンセル</button>
        <button disabled={!canSave} onClick={handleSave} className="px-4 py-1.5 text-sm rounded bg-[#1F2A3C] text-[#F3ECDA] disabled:opacity-40">{initial ? "更新" : "追加"}</button>
      </div>
    </div>
  );
}

function EnNoteInputs({ en, setEn, note, setNote }) {
  return (
    <div className="flex-1 flex flex-col gap-2">
      <input value={en} onChange={(e) => setEn(e.target.value)} placeholder="This is a book." className="rounded border border-[#c9bb9a] bg-white/60 px-3 py-2 font-display text-[#3B2A1C]" autoFocus />
      <input value={note} onChange={(e) => setNote(e.target.value)} placeholder="メモ（任意・出題には使われません）" className="rounded border border-[#c9bb9a] bg-white/40 px-3 py-1.5 text-sm text-[#6b5d44]" />
    </div>
  );
}

/* ---------------- Review Session ---------------- */
function buildQueue(lesson, progress) {
  const now = Date.now();
  const getProg = (c) => progress[c.id] || emptyProgress();
  const due = lesson.cards.filter((c) => (getProg(c).dueAt || 0) <= now);
  const pool = due.length > 0 ? due : [...lesson.cards].sort((a, b) => (getProg(a).lastReview || 0) - (getProg(b).lastReview || 0));
  const size = Math.min(10, Math.max(pool.length, Math.min(5, lesson.cards.length)));
  const shuffled = [...pool].sort(() => Math.random() - 0.5).slice(0, size);
  return shuffled.length > 0 ? shuffled : [...lesson.cards].sort(() => Math.random() - 0.5).slice(0, Math.min(10, lesson.cards.length));
}
function makeQuestion(card, allCards) {
  const type = Math.random() < 0.5 ? "pic2en" : "en2pic";
  const distractPool = allCards.filter((c) => c.id !== card.id && c.en !== card.en);
  const shuffled = [...distractPool].sort(() => Math.random() - 0.5).slice(0, 3);
  const choiceCards = [...shuffled, card].sort(() => Math.random() - 0.5);
  return { card, type, choices: choiceCards };
}

function ReviewSession({ lesson, allCards, initialProgress, onExit, onFinish }) {
  const queue = useRef(buildQueue(lesson, initialProgress));
  const [qIndex, setQIndex] = useState(0);
  const [question, setQuestion] = useState(() => makeQuestion(queue.current[0], allCards));
  const [selected, setSelected] = useState(null);
  const [correctCount, setCorrectCount] = useState(0);
  const [updates, setUpdates] = useState([]);
  const [xp, setXp] = useState(0);

  const total = queue.current.length;

  const handleChoice = (choice) => {
    if (selected) return;
    const isCorrect = choice.id === question.card.id;
    setSelected({ choiceId: choice.id, isCorrect });

    const now = Date.now();
    const prevLevel = (initialProgress[question.card.id] || emptyProgress()).level;
    const alreadyUpdated = updates.find((u) => u.id === question.card.id);
    const baseLevel = alreadyUpdated ? alreadyUpdated.level : prevLevel;
    const newLevel = isCorrect ? Math.min(baseLevel + 1, 5) : Math.max(baseLevel - 1, 0);
    const dueAt = now + INTERVAL_DAYS[newLevel] * DAY_MS;
    setUpdates((u) => [...u.filter((x) => x.id !== question.card.id), { id: question.card.id, level: newLevel, dueAt, lastReview: now }]);
    if (isCorrect) { setCorrectCount((n) => n + 1); setXp((x) => x + 10); }
  };

  const handleNext = () => {
    const nextIndex = qIndex + 1;
    if (nextIndex >= total) { onFinish({ total, correct: correctCount, xpEarned: xp, updates }); return; }
    setQIndex(nextIndex);
    setQuestion(makeQuestion(queue.current[nextIndex], allCards));
    setSelected(null);
  };

  const q = question;

  return (
    <div className="pt-6">
      <div className="flex items-center justify-between mb-4">
        <button onClick={onExit} className="text-[#D9C79A] text-sm flex items-center gap-1 hover:text-[#F3ECDA]"><X size={16} /> 終了</button>
        <div className="font-mono text-xs text-[#D9C79A]">{qIndex + 1} / {total}</div>
      </div>
      <div className="w-full h-2 bg-[#4a3520] rounded-full mb-6 overflow-hidden">
        <div className="h-full bg-[#E3B355] transition-all duration-300" style={{ width: `${((qIndex + (selected ? 1 : 0)) / total) * 100}%` }} />
      </div>

      <div className="card-paper rounded-md border border-[#c9bb9a] p-6 mb-5 text-center relative">
        <div className="punch-hole absolute -top-2 left-1/2 -translate-x-1/2" />
        {q.type === "pic2en" ? (
          <div className="flex items-center justify-center min-h-[110px]">
            <CardVisual card={q.card} className="max-h-32 max-w-[220px] object-contain rounded text-7xl" iconSize={64} />
          </div>
        ) : (
          <div className="font-display text-2xl font-bold text-[#3B2A1C] py-4 border-b-2 border-[#c9bb9a] inline-block px-4">{q.card.en}</div>
        )}
      </div>

      <div className={`grid gap-3 ${q.choices.length > 2 ? "grid-cols-2" : "grid-cols-1"}`}>
        {q.choices.map((choice) => {
          const isThisSelected = selected?.choiceId === choice.id;
          const isAnswer = choice.id === q.card.id;
          let cls = "bg-white/70 border-[#c9bb9a] text-[#3B2A1C] hover:bg-white";
          if (selected) {
            if (isAnswer) cls = "bg-[#dcefe2] border-[#2F6B4F] text-[#1e4632]";
            else if (isThisSelected) cls = "bg-[#f6e3dd] border-[#a3402f] text-[#7c2c1f]";
            else cls = "bg-white/30 border-[#c9bb9a] text-[#8a7a5c] opacity-60";
          }
          return (
            <button key={choice.id} onClick={() => handleChoice(choice)} disabled={!!selected} className={`rounded-md border-2 p-4 flex items-center justify-center gap-2 font-display font-semibold text-lg transition min-h-[64px] ${cls}`}>
              {selected && isAnswer && <Check size={18} className="stamp shrink-0" />}
              {selected && isThisSelected && !isAnswer && <X size={18} className="stamp shrink-0" />}
              {q.type === "pic2en" ? choice.en : <CardVisual card={choice} className="max-h-14 max-w-[90px] object-contain text-3xl" iconSize={28} />}
            </button>
          );
        })}
      </div>

      {selected && (
        <div className="mt-5 slide-up">
          <div className={`text-center font-display font-bold text-lg mb-3 ${selected.isCorrect ? "text-[#2F6B4F]" : "text-[#a3402f]"}`}>
            {selected.isCorrect ? "✓ 正解！" : q.type === "pic2en" ? `✗ ${q.card.en}` : "✗"}
          </div>
          <button onClick={handleNext} className="w-full rounded-md py-3 bg-[#1F2A3C] text-[#F3ECDA] font-display font-bold text-lg hover:brightness-110 transition">
            {qIndex + 1 >= total ? "結果を見る" : "次へ"}
          </button>
        </div>
      )}
    </div>
  );
}

/* ---------------- Summary ---------------- */
function Summary({ result, onReviewAgain, onBackToLesson, onHome }) {
  const pct = result.total > 0 ? Math.round((result.correct / result.total) * 100) : 0;
  return (
    <div className="pt-10 text-center slide-up">
      <div className="text-6xl mb-3">{pct >= 80 ? "🏅" : pct >= 50 ? "📗" : "📖"}</div>
      <h2 className="font-display text-2xl font-bold text-[#F3ECDA] mb-1">セッション終了</h2>
      <p className="text-[#D9C79A] mb-6 font-mono text-sm">{result.correct} / {result.total} 正解（{pct}%）・ +{result.xpEarned} XP</p>
      <div className="flex flex-col gap-3 max-w-xs mx-auto">
        <button onClick={onReviewAgain} className="rounded-md py-3 flex items-center justify-center gap-2 bg-[#2F6B4F] text-[#F3ECDA] font-display font-bold hover:brightness-110 transition"><RotateCcw size={18} /> もう一度復習する</button>
        <button onClick={onBackToLesson} className="rounded-md py-3 flex items-center justify-center gap-2 bg-[#5a4a30] text-[#F3ECDA] font-display font-semibold hover:brightness-110 transition"><BookOpen size={18} /> カード一覧に戻る</button>
        <button onClick={onHome} className="rounded-md py-2 flex items-center justify-center gap-2 text-[#D9C79A] hover:text-[#F3ECDA] transition"><HomeIcon size={16} /> 引き出し一覧へ</button>
      </div>
    </div>
  );
}
