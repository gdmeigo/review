import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ArrowLeft, Check, X, Play,
  Flame, Star, ChevronRight, BookOpen, Loader2, RotateCcw,
  Home as HomeIcon, Settings, ImageOff,
  FileSpreadsheet, Volume2
} from "lucide-react";

/* ---------------------------------------------------------
   GDM Picture Review — a card-catalog styled flashcard drill
   for reviewing Richards' Graded Direct Method / English
   Through Pictures lessons. Picture-first, no translation —
   the review screen itself carries no Japanese instructions.

   Content (lessons & cards) is SHARED across everyone using
   this artifact. Each person's own review progress and XP
   stay PRIVATE to them. Settings imports a CSV / Excel file —
   images are just URLs typed into a spreadsheet
   column, no upload needed.
--------------------------------------------------------- */

const FONT_IMPORT = `@import url('https://fonts.googleapis.com/css2?family=Fraunces:ital,opsz,wght@0,9..144,500;0,9..144,700;0,9..144,900;1,9..144,600&family=Inter:wght@400;500;600;700&family=IBM+Plex+Mono:wght@500;600&display=swap');`;

const INTERVAL_DAYS = [0, 1, 2, 4, 7, 14]; // by mastery level 0-5
const DAY_MS = 24 * 60 * 60 * 1000;
const todayStr = () => new Date().toISOString().slice(0, 10);
const emptyProgress = () => ({ level: 0, dueAt: 0, lastReview: 0 });
const DRIVE_IMAGE_SIZE = "w1000";
const DEFAULT_SHEET_URL = "https://docs.google.com/spreadsheets/d/1NbV4QywhVxkT8iOs11CSM-12llyQhrtdVj0Gvs8dO84/edit?usp=drive_link";
const DEFAULT_SHEET_DOWNLOAD_URL = normalizeSheetUrl(DEFAULT_SHEET_URL);

function getGoogleDriveFileId(url) {
  try {
    const parsed = new URL(url);
    const host = parsed.hostname.toLowerCase();
    if (host === "lh3.googleusercontent.com") {
      const directMatch = parsed.pathname.match(/^\/d\/([^/=]+)/);
      return directMatch ? directMatch[1] : "";
    }
    if (host !== "drive.google.com" && host !== "docs.google.com" && host !== "drive.usercontent.google.com") return "";

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
  try {
    const parsed = new URL(trimmed);
    const sheetMatch = parsed.pathname.match(/\/spreadsheets\/d\/([^/]+)/);
    if (parsed.hostname.toLowerCase() === "docs.google.com" && sheetMatch) {
      return `https://docs.google.com/spreadsheets/d/${encodeURIComponent(sheetMatch[1])}/export?format=xlsx`;
    }
  } catch {
    // Fall through to the Drive file normalization below.
  }

  const driveFileId = getGoogleDriveFileId(trimmed);
  if (driveFileId) {
    return `https://drive.usercontent.google.com/download?id=${encodeURIComponent(driveFileId)}&export=download`;
  }

  return trimmed;
}

/* ---------------- spreadsheet import helpers ---------------- */
const FIELD_ALIASES = {
  lesson: ["lesson", "レッスン", "レッスン名", "unit", "ユニット"],
  item: ["item", "no", "number", "question", "worksheet", "work", "問題", "番号", "項目"],
  emoji: ["emoji", "絵文字"],
  en: ["english", "en", "英語", "英文", "word", "text", "単語"],
  sentence: ["sentence", "prompt", "line", "英語文", "文", "設問", "文章"],
  choices: ["choices", "options", "choice", "選択肢", "候補"],
  answer: ["answer", "correct", "正解", "答え"],
  image: ["image", "imageurl", "image_url", "画像", "画像url", "picture", "photo"],
  audio: ["audio", "audiourl", "audio_url", "sound", "音声", "音声url"],
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
function normalizeAudioUrl(url) {
  const trimmed = (url || "").trim();
  if (!trimmed) return "";
  const driveFileId = getGoogleDriveFileId(trimmed);
  if (driveFileId) {
    return `https://drive.google.com/file/d/${encodeURIComponent(driveFileId)}/preview`;
  }
  return trimmed;
}
function getGoogleDrivePreviewUrl(url) {
  const driveFileId = getGoogleDriveFileId(url);
  return driveFileId ? `https://drive.google.com/file/d/${encodeURIComponent(driveFileId)}/preview` : "";
}
function splitChoices(value) {
  return (value || "")
    .split(/[|／/、,，;]/)
    .map((choice) => choice.trim())
    .filter(Boolean);
}
function shuffleList(list) {
  return [...list].sort(() => Math.random() - 0.5);
}
function parseClozeSentence(value, answerValue) {
  const raw = value || "";
  const answers = [];
  const parts = [];
  const pattern = /\{\{([^}]+)\}\}/g;
  let lastIndex = 0;
  let match;
  while ((match = pattern.exec(raw))) {
    parts.push(raw.slice(lastIndex, match.index));
    answers.push(match[1].trim());
    lastIndex = match.index + match[0].length;
  }
  parts.push(raw.slice(lastIndex));
  const explicitAnswers = splitChoices(answerValue);
  if (answers.length === 0 && explicitAnswers.length > 0 && /_{2,}/.test(raw)) {
    return {
      sentence: raw,
      clozeParts: raw.split(/_{2,}/),
      answers: explicitAnswers,
    };
  }
  return {
    sentence: answers.length > 0 ? parts.join("____") : raw,
    clozeParts: answers.length > 0 ? parts : null,
    answers: explicitAnswers.length > 0 ? explicitAnswers : answers,
  };
}
function buildContentFromRows(rows) {
  const order = [];
  const byTitle = {};
  rows.forEach((row) => {
    const sentence = getField(row, "sentence");
    const en = sentence || getField(row, "en");
    const choices = splitChoices(getField(row, "choices"));
    const answer = getField(row, "answer");
    const cloze = parseClozeSentence(en, answer);
    const audioUrl = normalizeAudioUrl(getField(row, "audio"));
    if (!en && choices.length === 0 && !answer) return;
    const title = getField(row, "lesson") || "未分類";
    const item = getField(row, "item");
    const emoji = getField(row, "emoji");
    const image = normalizeImageUrl(getField(row, "image"));
    const note = getField(row, "note");
    if (!byTitle[title]) {
      byTitle[title] = { title, emoji: emoji || "📇", cards: [], cardsByItem: {} };
      order.push(title);
    } else if (emoji && byTitle[title].emoji === "📇") {
      byTitle[title].emoji = emoji;
    }
    const isWorksheetRow = !!(sentence || choices.length > 0 || answer || audioUrl || item);
    if (isWorksheetRow) {
      const itemKey = item || image || en || `row-${byTitle[title].cards.length + 1}`;
      let card = byTitle[title].cardsByItem[itemKey];
      if (!card) {
        card = { en: en || answer || itemKey, image, emoji, note, audioUrl, worksheetLines: [] };
        byTitle[title].cardsByItem[itemKey] = card;
        byTitle[title].cards.push(card);
      }
      if (image && !card.image) card.image = image;
      if (emoji && !card.emoji) card.emoji = emoji;
      if (note && !card.note) card.note = note;
      if (audioUrl && !card.audioUrl) card.audioUrl = audioUrl;
      card.worksheetLines.push({
        sentence: cloze.sentence,
        clozeParts: cloze.clozeParts,
        choices,
        answers: cloze.answers,
        audioUrl,
        note,
      });
    } else {
      byTitle[title].cards.push({ en, image, emoji, note, audioUrl });
    }
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
        audioUrl: c.audioUrl || undefined,
        worksheetLines: c.worksheetLines || undefined,
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

function parseSheetBuffer(buffer, sourceName = "", contentType = "") {
  const lower = sourceName.toLowerCase().split("?")[0];
  const type = (contentType || "").toLowerCase();
  const looksLikeExcel =
    lower.endsWith(".xlsx") ||
    lower.endsWith(".xls") ||
    type.includes("spreadsheet") ||
    type.includes("excel") ||
    isExcelBuffer(buffer);

  if (looksLikeExcel) {
    const wb = XLSX.read(buffer, { type: "array" });
    const sheet = wb.Sheets[wb.SheetNames[0]];
    return XLSX.utils.sheet_to_json(sheet, { defval: "" });
  }

  const text = new TextDecoder("utf-8").decode(buffer).replace(/^\uFEFF/, "");
  const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
  if (parsed.errors.length > 0) {
    throw new Error(parsed.errors[0].message || "CSV parse failed");
  }
  return parsed.data;
}

export default function App() {
  const [ready, setReady] = useState(false);
  const [lessons, setLessons] = useState({});
  const [index, setIndex] = useState([]);
  const [stats, setStats] = useState({ xp: 0, reviewDates: [] });
  const [progressByLesson, setProgressByLesson] = useState({});
  const [screen, setScreen] = useState({ name: "home" });

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

      const loaded = {};
      for (const meta of idx) {
        const raw = await safeGet("lesson:" + meta.id, true);
        if (raw) loaded[meta.id] = JSON.parse(raw);
      }
      setIndex(idx);
      setLessons(loaded);
      setStats(st);
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

  // ---- spreadsheet sync ----
  const importRows = async (rows) => {
    const { index: newIndex, lessonsMap: newLessons } = buildContentFromRows(rows);
    if (newIndex.length === 0) {
      throw new Error("CSVに english 列のカードが見つかりませんでした");
    }
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
          background-color: #f7fcff;
          background-image:
            radial-gradient(circle at 20% 10%, rgba(115,191,215,0.18), transparent 28%),
            linear-gradient(180deg, #ffffff 0%, #edf9fc 55%, #f7fcff 100%);
        }
        .card-paper {
          background: rgba(255,255,255,0.92);
          background-image: radial-gradient(rgba(22,135,167,0.06) 1px, transparent 1px);
          background-size: 6px 6px;
        }
        .punch-hole { width: 14px; height: 14px; border-radius: 50%; background: #d7eef6; box-shadow: inset 0 1px 2px rgba(22,71,95,0.22); }
        .drawer-front { background: linear-gradient(180deg, #ffffff 0%, #e8f7fb 100%); border: 1px solid #b7d6e6; }
        .brass { background: linear-gradient(180deg, #d7f5eb 0%, #9fe3d1 100%); border: 1px solid #73cdb9; }
        @keyframes stampIn { 0% { transform: scale(2.2) rotate(-12deg); opacity: 0; } 60% { transform: scale(0.95) rotate(-12deg); opacity: 1; } 100% { transform: scale(1) rotate(-12deg); opacity: 1; } }
        .stamp { animation: stampIn 0.35s ease-out; }
        @keyframes slideUp { from { transform: translateY(14px); opacity: 0; } to { transform: translateY(0); opacity: 1; } }
        .slide-up { animation: slideUp 0.28s ease-out; }
        @media (prefers-reduced-motion: reduce) { .stamp, .slide-up { animation: none; } }
        button:focus-visible, input:focus-visible, textarea:focus-visible { outline: 3px solid #73bfd7; outline-offset: 2px; }
      `}</style>

      {!ready ? (
        <div className="cabinet-bg min-h-screen flex items-center justify-center">
          <Loader2 className="animate-spin text-[#73bfd7]" size={32} />
        </div>
      ) : (
        <div className="cabinet-bg min-h-screen">
          <TopBar screen={screen} setScreen={setScreen} xp={stats.xp} streak={streak} />
          <div className="max-w-2xl mx-auto px-4 pb-16">
            {screen.name === "home" && (
              <Home
                index={index}
                onOpen={(id) => setScreen({ name: "lesson", id })}
              />
            )}
            {screen.name === "settings" && (
              <SettingsScreen onImportRows={importRows} />
            )}
            {screen.name === "lesson" && lessons[screen.id] && (
              <LessonDetail
                lesson={lessons[screen.id]}
                progress={progressByLesson[screen.id]}
                onEnsureProgress={() => loadProgress(screen.id)}
                onBack={() => setScreen({ name: "home" })}
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
    </div>
  );
}

/* ---------------- TopBar ---------------- */
function TopBar({ screen, setScreen, xp, streak }) {
  return (
    <div className="sticky top-0 z-30 backdrop-blur-sm bg-white/90 border-b border-[#d7eef6]">
      <div className="max-w-2xl mx-auto px-4 py-3 flex items-center justify-between gap-2">
        <button onClick={() => setScreen({ name: "home" })} className="flex items-center gap-2 text-[#16475f] min-w-0">
          <span className="font-display text-lg font-bold tracking-tight truncate">GDM 復習カード</span>
        </button>
        <div className="flex items-center gap-3 font-mono text-sm shrink-0">
          <div className="flex items-center gap-1 text-[#1687a7]"><Flame size={16} /> {streak}</div>
          <div className="flex items-center gap-1 text-[#16805d]"><Star size={16} /> {xp}</div>
          <button onClick={() => setScreen({ name: "settings" })} className="flex items-center gap-1 text-[10px] bg-[#1687a7] text-white px-2 py-1 rounded" title="設定">
            <Settings size={12} /> 設定
          </button>
          {screen.name !== "home" && (
            <button
              onClick={() => setScreen(screen.name === "summary" ? { name: "home" } : screen.name === "review" ? { name: "lesson", id: screen.id } : { name: "home" })}
              className="text-[#1687a7] hover:text-[#16475f]"
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

/* ---------------- Home ---------------- */
function Home({ index, onOpen }) {
  return (
    <div className="pt-6">
      <p className="font-display italic text-[#1687a7] mb-2 text-[15px]">Look at the picture. Say it in English.</p>
      <p className="text-[11px] text-[#42677a] mb-5 font-mono">レッスン内容は共有されます。復習の記録(習熟度・XP)は自分だけに保存されます。</p>

      {index.length === 0 && (
        <div className="card-paper rounded-md p-6 text-center mb-4 border border-[#b7d6e6]">
          <p className="font-display text-lg text-[#16475f] mb-1">まだレッスンがありません</p>
          <p className="text-sm text-[#42677a]">右上の設定からCSV/Excelを読み込んでください</p>
        </div>
      )}

      <div className="space-y-3">
        {index.map((meta) => (
          <button key={meta.id} onClick={() => onOpen(meta.id)} className="drawer-front w-full rounded-md p-4 flex items-center gap-4 text-left shadow-md hover:brightness-110 transition">
            <div className="brass rounded w-12 h-12 flex items-center justify-center text-2xl shrink-0">{meta.emoji || "📇"}</div>
            <div className="flex-1 min-w-0">
              <div className="font-display text-[#16475f] text-lg font-semibold truncate">{meta.title}</div>
              <div className="font-mono text-xs text-[#42677a]">{meta.count || 0} 枚のカード</div>
            </div>
            <ChevronRight className="text-[#1687a7]" size={20} />
          </button>
        ))}
      </div>
    </div>
  );
}

function SettingsScreen({ onImportRows }) {
  return (
    <div className="pt-6">
      <SheetSyncPanel onImportRows={onImportRows} />
    </div>
  );
}

/* ---------------- Settings / Import Panel ---------------- */
function SheetSyncPanel({ onImportRows }) {
  const [status, setStatus] = useState(null);
  const [showHelp, setShowHelp] = useState(false);
  const csvExample = `lesson,item,image,sentence,choices,answer,audio,note
WorkP22,1,https://example.com/picture1.png,This is a {{man}}.,man|woman,,https://example.com/audio1.mp3,{{ }} の中が空欄になります
WorkP22,1,https://example.com/picture1.png,Her {{book}} is on the {{table}}.,book|bag|table|desk,,,
WorkP22,2,https://example.com/picture2.png,This is a ____.,glass|bottle,glass,,____ を使う場合は answer に正解を書きます`;

  const doFileImport = async (file) => {
    if (!file) return;
    setStatus({ type: "loading", message: "読み込み中..." });
    try {
      const rows = parseSheetBuffer(await file.arrayBuffer(), file.name, file.type);
      const newIndex = await onImportRows(rows);
      const totalCards = newIndex.reduce((s, m) => s + (m.count || 0), 0);
      setStatus({ type: "success", message: newIndex.length + "レッスン・" + totalCards + "枚のカードを読み込みました" });
    } catch (e) {
      setStatus({ type: "error", message: "読み込みに失敗しました: " + (e?.message || "CSVまたはExcelファイルを確認してください") });
    }
  };

  return (
    <div className="card-paper rounded-md p-4 mt-4 border border-[#b7d6e6]">
      <div className="flex items-center gap-2 mb-2 text-[#16475f]">
        <FileSpreadsheet size={18} />
        <h3 className="font-display font-bold">設定</h3>
      </div>
      <p className="text-xs text-[#42677a] mb-3">
        CSVまたはExcelファイルからワークシートを読み込めます。列は
        <span className="font-mono"> lesson / item / image / sentence / choices / answer / audio / note </span>
        を使えます。画像と音声はURLで指定できます。Google Driveの音声はリンクを知っている人が閲覧できる設定にし、アプリ内ではDriveのプレビュープレイヤーで再生します。
      </p>
      <button onClick={() => setShowHelp((s) => !s)} className="text-xs text-[#1687a7] underline mb-3">
        {showHelp ? "手順を隠す" : "読み込み手順を見る"}
      </button>
      {showHelp && (
        <ol className="text-xs text-[#42677a] list-decimal list-inside mb-3 space-y-1">
          <li>下の既定Excelファイルをダウンロードするか、自分のCSV/Excelファイルを用意</li>
          <li>「CSV/Excelファイルを選んで読み込む」からファイルを選択</li>
          <li>同じ lesson と item の行は、1つの絵に複数の文があるワークシート項目としてまとまります</li>
          <li>読み込むと現在のレッスン内容がファイルの内容に置き換わります</li>
        </ol>
      )}
      <div className="mb-3 rounded border border-[#b7d6e6] bg-white/70 p-3 text-xs text-[#42677a]">
        <div className="mb-2 font-display font-bold text-[#16475f]">CSV記載例</div>
        <p className="mb-2">
          <span className="font-mono">choices</span> は1つの列にまとめ、
          <span className="font-mono"> man|woman|girl </span>
          のように区切って書きます。問題では選択肢がランダム順で表示されます。
        </p>
        <pre className="overflow-x-auto whitespace-pre rounded bg-[#f4fbfd] p-2 font-mono text-[10px] leading-5 text-[#16475f]">{csvExample}</pre>
      </div>
      <a
        href={DEFAULT_SHEET_DOWNLOAD_URL}
        target="_blank"
        rel="noreferrer"
        className="mb-2 w-full rounded-md py-2 flex items-center justify-center gap-2 border border-[#73bfd7] text-[#166078] font-display font-semibold hover:bg-[#e8f7fb]"
      >
        <FileSpreadsheet size={16} />
        既定Excelファイルをダウンロードする
      </a>
      <label className="mt-2 w-full rounded-md py-2 flex items-center justify-center gap-2 border border-[#73bfd7] text-[#166078] font-display font-semibold cursor-pointer hover:bg-[#e8f7fb]">
        <FileSpreadsheet size={16} />
        CSV/Excelファイルを選んで読み込む
        <input
          type="file"
          accept=".csv,.tsv,.txt,.xlsx,.xls,text/csv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel"
          className="hidden"
          onChange={(e) => {
            doFileImport(e.target.files?.[0]);
            e.target.value = "";
          }}
        />
      </label>
      {status && status.type !== "loading" && (
        <p className={"text-xs mt-2 " + (status.type === "success" ? "text-[#16805d]" : "text-[#b42335]")}>{status.message}</p>
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

function AudioButton({ src, label = "音声" }) {
  const audioRef = useRef(null);
  const [isPlaying, setIsPlaying] = useState(false);
  const [hasError, setHasError] = useState(false);
  const drivePreviewUrl = getGoogleDrivePreviewUrl(src);
  const playableSrc = normalizeAudioUrl(src);

  useEffect(() => {
    return () => {
      if (audioRef.current) {
        audioRef.current.pause();
        audioRef.current = null;
      }
    };
  }, []);

  if (!playableSrc) return null;

  if (drivePreviewUrl) {
    return (
      <span className="inline-flex max-w-full flex-col items-start gap-1">
        <span className="inline-flex items-center gap-1 text-[11px] text-[#166078]">
          <Volume2 size={13} />
          {label}
        </span>
        <iframe
          src={drivePreviewUrl}
          title={label}
          allow="autoplay"
          className="h-20 w-64 max-w-full rounded border border-[#b7d6e6] bg-white"
        />
      </span>
    );
  }

  const handlePlay = async () => {
    setHasError(false);
    if (!audioRef.current || audioRef.current.src !== playableSrc) {
      if (audioRef.current) audioRef.current.pause();
      audioRef.current = new Audio(playableSrc);
      audioRef.current.addEventListener("ended", () => setIsPlaying(false));
      audioRef.current.addEventListener("pause", () => setIsPlaying(false));
      audioRef.current.addEventListener("play", () => setIsPlaying(true));
      audioRef.current.addEventListener("error", () => {
        setIsPlaying(false);
        setHasError(true);
      });
    }

    try {
      if (isPlaying) {
        audioRef.current.pause();
      } else {
        await audioRef.current.play();
      }
    } catch {
      setIsPlaying(false);
      setHasError(true);
    }
  };

  return (
    <button
      type="button"
      onClick={handlePlay}
      className="inline-flex items-center gap-1 rounded border border-[#73bfd7] px-2 py-1 text-[11px] text-[#166078] hover:bg-[#e8f7fb]"
      title={hasError ? "音声を再生できませんでした" : label}
    >
      <Volume2 size={13} />
      {hasError ? "再生不可" : isPlaying ? "停止" : label}
    </button>
  );
}

function WorksheetLines({ lines, compact = false }) {
  if (!lines?.length) return null;
  return (
    <div className={compact ? "space-y-1" : "space-y-2"}>
      {lines.map((line, index) => (
        <div key={index} className="text-left">
          <div className={compact ? "text-xs text-[#16475f]" : "text-sm text-[#16475f]"}>
            {line.clozeParts ? (
              line.clozeParts.map((part, partIndex) => (
                <span key={partIndex}>
                  {part}
                  {partIndex < line.clozeParts.length - 1 && <span className="inline-block min-w-12 border-b border-[#16475f] align-baseline" />}
                </span>
              ))
            ) : (
              line.sentence || line.answers?.join(" / ")
            )}
          </div>
          {line.choices?.length > 0 && (
            <div className="mt-1 flex flex-wrap gap-1">
              {line.choices.map((choice) => (
                <span key={choice} className="rounded border border-[#b7d6e6] bg-white/70 px-2 py-0.5 text-[11px] text-[#42677a]">
                  {choice}
                </span>
              ))}
            </div>
          )}
          {line.audioUrl && <div className="mt-1"><AudioButton src={line.audioUrl} /></div>}
        </div>
      ))}
    </div>
  );
}

function WorksheetQuestion({ card, selected, onSubmit }) {
  const [fills, setFills] = useState({});
  const lines = card.worksheetLines || [];

  useEffect(() => {
    setFills({});
  }, [card.id]);

  const shuffledChoicesByLine = useMemo(
    () => lines.map((line) => shuffleList(line.choices || [])),
    [card.id, lines]
  );
  const allFilled = lines.every((line, lineIndex) =>
    (line.answers || []).every((_, blankIndex) => fills[`${lineIndex}-${blankIndex}`])
  );

  const setBlank = (lineIndex, blankIndex, choice) => {
    if (selected) return;
    setFills((prev) => ({ ...prev, [`${lineIndex}-${blankIndex}`]: choice }));
  };

  const isCorrect = () =>
    lines.every((line, lineIndex) =>
      (line.answers || []).every((answer, blankIndex) => fills[`${lineIndex}-${blankIndex}`] === answer)
    );

  return (
    <div className="mx-auto max-w-md py-2 text-left">
      {card.photoUrl && (
        <div className="mb-4 flex justify-center">
          <CardVisual card={card} className="max-h-44 max-w-full object-contain rounded" iconSize={48} />
        </div>
      )}
      <div className="space-y-4">
        {lines.map((line, lineIndex) => (
          <div key={lineIndex}>
            <div className="text-sm leading-8 text-[#16475f]">
              {(line.clozeParts || [line.sentence || ""]).map((part, partIndex) => (
                <span key={partIndex}>
                  {part}
                  {partIndex < (line.answers || []).length && (
                    <span className={`mx-1 inline-flex min-w-20 items-center justify-center border-b-2 px-2 ${selected ? fills[`${lineIndex}-${partIndex}`] === line.answers[partIndex] ? "border-[#16805d] text-[#16805d]" : "border-[#b42335] text-[#b42335]" : "border-[#1687a7] text-[#16475f]"}`}>
                      {fills[`${lineIndex}-${partIndex}`] || " "}
                    </span>
                  )}
                </span>
              ))}
            </div>
            <div className="mt-2 flex flex-wrap gap-2">
              {(shuffledChoicesByLine[lineIndex] || []).map((choice) => (
                <button
                  key={choice}
                  type="button"
                  disabled={!!selected}
                  onClick={() => {
                    const nextBlank = (line.answers || []).findIndex((_, blankIndex) => !fills[`${lineIndex}-${blankIndex}`]);
                    if (nextBlank >= 0) setBlank(lineIndex, nextBlank, choice);
                  }}
                  className="rounded border border-[#b7d6e6] bg-white px-2 py-1 text-xs text-[#166078] hover:bg-[#e8f7fb] disabled:opacity-70"
                >
                  {choice}
                </button>
              ))}
              {(line.answers || []).some((_, blankIndex) => fills[`${lineIndex}-${blankIndex}`]) && !selected && (
                <button
                  type="button"
                  onClick={() => setFills((prev) => {
                    const next = { ...prev };
                    (line.answers || []).forEach((_, blankIndex) => delete next[`${lineIndex}-${blankIndex}`]);
                    return next;
                  })}
                  className="rounded px-2 py-1 text-xs text-[#42677a] underline"
                >
                  クリア
                </button>
              )}
            </div>
            {line.audioUrl && <div className="mt-2"><AudioButton src={line.audioUrl} /></div>}
          </div>
        ))}
      </div>
      {!selected && (
        <button
          type="button"
          disabled={!allFilled}
          onClick={() => onSubmit(isCorrect())}
          className="mt-5 w-full rounded-md bg-[#1687a7] py-2 font-display font-bold text-white disabled:opacity-40"
        >
          答える
        </button>
      )}
    </div>
  );
}

/* ---------------- Lesson Detail ---------------- */
function LessonDetail({ lesson, progress, onEnsureProgress, onBack, onStartReview }) {
  useEffect(() => {
    onEnsureProgress();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lesson.id]);

  const prog = progress || {};

  return (
    <div className="pt-6">
      <button onClick={onBack} className="flex items-center gap-1 text-[#1687a7] text-sm mb-4 hover:text-[#16475f]">
        <ArrowLeft size={16} /> 引き出し一覧へ
      </button>

      <div className="card-paper rounded-md p-4 border border-[#b7d6e6] mb-4">
        <div className="flex items-center gap-3">
          <div className="brass rounded w-11 h-11 flex items-center justify-center text-xl shrink-0">{lesson.emoji}</div>
          <div className="flex-1">
            <div className="font-display text-xl font-bold text-[#16475f]">{lesson.title}</div>
            <div className="font-mono text-xs text-[#42677a]">{lesson.cards.length} 枚</div>
          </div>
        </div>
      </div>

      <button onClick={onStartReview} disabled={lesson.cards.length < 1} className="w-full mb-5 rounded-md py-3 flex items-center justify-center gap-2 bg-[#16805d] text-[#ffffff] font-display font-bold text-lg shadow-md disabled:opacity-40 disabled:cursor-not-allowed hover:brightness-110 transition">
        <Play size={20} /> 復習をはじめる
      </button>
      {lesson.cards.length < 1 && <p className="text-xs text-[#1687a7] -mt-3 mb-4">復習を始めるにはカードが必要です</p>}

      <div className="space-y-2 mb-4">
        {lesson.cards.map((c) => (
          <div key={c.id} className={`card-paper rounded-md p-3 border border-[#b7d6e6] flex gap-3 ${c.worksheetLines?.length ? "items-start" : "items-center"}`}>
            <div className="punch-hole ml-1 hidden sm:block" />
            <div className={`${c.worksheetLines?.length ? "w-24 h-24" : "w-10 h-10"} shrink-0 flex items-center justify-center text-2xl rounded overflow-hidden bg-white/40`}>
              <CardVisual card={c} className={`${c.worksheetLines?.length ? "w-24 h-24" : "w-10 h-10"} object-cover rounded text-2xl flex items-center justify-center`} iconSize={18} />
            </div>
            <div className="flex-1 min-w-0">
              {c.worksheetLines?.length ? (
                <WorksheetLines lines={c.worksheetLines} compact />
              ) : (
                <>
                  <div className="font-display text-[#16475f] font-semibold truncate">{c.en}</div>
                  {c.note && <div className="text-xs text-[#42677a] truncate">{c.note}</div>}
                  {c.audioUrl && <div className="mt-1"><AudioButton src={c.audioUrl} /></div>}
                </>
              )}
            </div>
            <MasteryDots level={(prog[c.id] || emptyProgress()).level} />
          </div>
        ))}
      </div>
    </div>
  );
}

function MasteryDots({ level }) {
  return (
    <div className="hidden sm:flex gap-0.5 mr-1" aria-label={`習熟度 ${level}/5`}>
      {[0, 1, 2, 3, 4].map((i) => <div key={i} className={`w-1.5 h-1.5 rounded-full ${i < level ? "bg-[#16805d]" : "bg-[#b7d6e6]"}`} />)}
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
  if (card.worksheetLines?.length) return { card, type: "worksheet", choices: [] };
  const type = Math.random() < 0.5 ? "pic2en" : "en2pic";
  const distractPool = allCards.filter((c) => c.id !== card.id && c.en !== card.en);
  const shuffled = [...distractPool].sort(() => Math.random() - 0.5).slice(0, 3);
  const choiceCards = [...shuffled, card].sort(() => Math.random() - 0.5);
  return { card, type, choices: choiceCards };
}
function cardLabel(card) {
  return card.worksheetLines?.[0]?.sentence || card.en;
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

  const recordAnswer = (isCorrect, choiceId = question.card.id) => {
    if (selected) return;
    setSelected({ choiceId, isCorrect });

    const now = Date.now();
    const prevLevel = (initialProgress[question.card.id] || emptyProgress()).level;
    const alreadyUpdated = updates.find((u) => u.id === question.card.id);
    const baseLevel = alreadyUpdated ? alreadyUpdated.level : prevLevel;
    const newLevel = isCorrect ? Math.min(baseLevel + 1, 5) : Math.max(baseLevel - 1, 0);
    const dueAt = now + INTERVAL_DAYS[newLevel] * DAY_MS;
    setUpdates((u) => [...u.filter((x) => x.id !== question.card.id), { id: question.card.id, level: newLevel, dueAt, lastReview: now }]);
    if (isCorrect) { setCorrectCount((n) => n + 1); setXp((x) => x + 10); }
  };
  const handleChoice = (choice) => {
    recordAnswer(choice.id === question.card.id, choice.id);
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
        <button onClick={onExit} className="text-[#1687a7] text-sm flex items-center gap-1 hover:text-[#16475f]"><X size={16} /> 終了</button>
        <div className="font-mono text-xs text-[#1687a7]">{qIndex + 1} / {total}</div>
      </div>
      <div className="w-full h-2 bg-[#d7eef6] rounded-full mb-6 overflow-hidden">
        <div className="h-full bg-[#73bfd7] transition-all duration-300" style={{ width: `${((qIndex + (selected ? 1 : 0)) / total) * 100}%` }} />
      </div>

      <div className="card-paper rounded-md border border-[#b7d6e6] p-6 mb-5 text-center relative">
        <div className="punch-hole absolute -top-2 left-1/2 -translate-x-1/2" />
        {q.type === "worksheet" ? (
          <WorksheetQuestion key={q.card.id} card={q.card} selected={selected} onSubmit={(isCorrect) => recordAnswer(isCorrect)} />
        ) : q.type === "pic2en" ? (
          <div className="flex items-center justify-center min-h-[110px]">
            <CardVisual card={q.card} className="max-h-32 max-w-[220px] object-contain rounded text-7xl" iconSize={64} />
          </div>
        ) : (
          <div className="font-display text-2xl font-bold text-[#16475f] py-4 border-b-2 border-[#b7d6e6] inline-block px-4">{q.card.en}</div>
        )}
      </div>

      {q.type !== "worksheet" && (
        <div className={`grid gap-3 ${q.choices.length > 2 ? "grid-cols-2" : "grid-cols-1"}`}>
          {q.choices.map((choice) => {
            const isThisSelected = selected?.choiceId === choice.id;
            const isAnswer = choice.id === q.card.id;
            let cls = "bg-white/70 border-[#b7d6e6] text-[#16475f] hover:bg-white";
            if (selected) {
              if (isAnswer) cls = "bg-[#dcefe2] border-[#16805d] text-[#1e4632]";
              else if (isThisSelected) cls = "bg-[#fff1f3] border-[#b42335] text-[#b42335]";
              else cls = "bg-white/30 border-[#b7d6e6] text-[#6b8794] opacity-60";
            }
            return (
              <button key={choice.id} onClick={() => handleChoice(choice)} disabled={!!selected} className={`rounded-md border-2 p-4 flex items-center justify-center gap-2 font-display font-semibold text-lg transition min-h-[64px] ${cls}`}>
                {selected && isAnswer && <Check size={18} className="stamp shrink-0" />}
                {selected && isThisSelected && !isAnswer && <X size={18} className="stamp shrink-0" />}
                {q.type === "pic2en" ? cardLabel(choice) : <CardVisual card={choice} className="max-h-14 max-w-[90px] object-contain text-3xl" iconSize={28} />}
              </button>
            );
          })}
        </div>
      )}

      {selected && (
        <div className="mt-5 slide-up">
          <div className={`text-center font-display font-bold text-lg mb-3 ${selected.isCorrect ? "text-[#16805d]" : "text-[#b42335]"}`}>
            {selected.isCorrect ? "✓ 正解！" : q.type === "pic2en" ? `✗ ${q.card.en}` : "✗"}
          </div>
          <button onClick={handleNext} className="w-full rounded-md py-3 bg-[#1687a7] text-[#ffffff] font-display font-bold text-lg hover:brightness-110 transition">
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
      <h2 className="font-display text-2xl font-bold text-[#16475f] mb-1">セッション終了</h2>
      <p className="text-[#1687a7] mb-6 font-mono text-sm">{result.correct} / {result.total} 正解（{pct}%）・ +{result.xpEarned} XP</p>
      <div className="flex flex-col gap-3 max-w-xs mx-auto">
        <button onClick={onReviewAgain} className="rounded-md py-3 flex items-center justify-center gap-2 bg-[#16805d] text-[#ffffff] font-display font-bold hover:brightness-110 transition"><RotateCcw size={18} /> もう一度復習する</button>
        <button onClick={onBackToLesson} className="rounded-md py-3 flex items-center justify-center gap-2 bg-[#1687a7] text-[#ffffff] font-display font-semibold hover:brightness-110 transition"><BookOpen size={18} /> カード一覧に戻る</button>
        <button onClick={onHome} className="rounded-md py-2 flex items-center justify-center gap-2 text-[#1687a7] hover:text-[#16475f] transition"><HomeIcon size={16} /> 引き出し一覧へ</button>
      </div>
    </div>
  );
}
