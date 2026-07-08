import { useState, useEffect, useCallback, useMemo, useRef } from "react";
import Papa from "papaparse";
import * as XLSX from "xlsx";
import {
  ArrowLeft, Check, X, Play,
  Flame, Star, ChevronRight, Loader2, RotateCcw,
  Home as HomeIcon, Settings, ImageOff,
  FileSpreadsheet, Volume2, Download, Upload, Palette
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

const MAX_LEVEL = 9999;
const LEVEL_GAIN_CORRECT = 30;
const LEVEL_LOSS_WRONG = 8;
const XP_PER_PERSONAL_LEVEL = 20;
const PERFECT_BADGE_LEVEL_BONUS = 25;
const XP_LEVEL_THRESHOLD = 100;
const PERFECT_BADGE_XP_BONUS = 30;
const EXTRA_PERFECT_BADGE_XP_BONUS = 10;
const DAY_MS = 24 * 60 * 60 * 1000;
const todayStr = () => new Date().toISOString().slice(0, 10);
const emptyProgress = () => ({ level: 0, dueAt: 0, lastReview: 0 });
const DRIVE_IMAGE_SIZE = "w1000";
const DEFAULT_EXCEL_URL = "https://docs.google.com/spreadsheets/d/1NbV4QywhVxkT8iOs11CSM-12llyQhrtdVj0Gvs8dO84/export?format=xlsx";
const DEFAULT_SHEET_DOWNLOAD_URL = DEFAULT_EXCEL_URL;
const CONTENT_IMPORTED_KEY = "content-imported";
const CONTENT_VERSION = "2026-07-05-google-sheet-default";
const USER_ID_KEY = "viewer-id";
const USER_PREFS_KEY = "user-prefs";
const DEFAULT_USER_PREFS = { tone: "fresh" };
const APP_NAME = "GDM Review";
const APP_VERSION = "1.0.0";
const PERSONAL_EXPORT_TYPE = "gdm-review-personal-settings";
const PERSONAL_EXPORT_VERSION = 2;

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
  lessonNo: ["lesson_no", "lessonno", "lesson_number", "レッスン番号"],
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
  hint: ["hint", "ヒント"],
  hintImage: ["hint_image", "hintimage", "hint_image_url", "hintimageurl", "ヒント画像", "ヒント画像url"],
  point: ["point", "ポイント"],
  pointImage: ["point_image", "pointimage", "point_image_url", "pointimageurl", "ポイント画像", "ポイント画像url"],
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
    const clozeParts = raw.split(/_{2,}/);
    return {
      sentence: raw,
      clozeParts,
      answers: explicitAnswers,
      blankCount: clozeParts.length - 1,
    };
  }
  if (answers.length === 0 && /_{2,}/.test(raw)) {
    const clozeParts = raw.split(/_{2,}/);
    const blankCount = clozeParts.length - 1;
    return {
      sentence: raw,
      clozeParts,
      answers: Array.from({ length: blankCount }, () => ""),
      blankCount,
    };
  }
  return {
    sentence: answers.length > 0 ? parts.join("____") : raw,
    clozeParts: answers.length > 0 ? parts : null,
    answers: explicitAnswers.length > 0 ? explicitAnswers : answers,
    blankCount: explicitAnswers.length > 0 ? explicitAnswers.length : answers.length,
  };
}
function isAnswerCorrect(expected, actual) {
  return !expected || actual === expected;
}
function fillClozeText(line, values) {
  const parts = line?.clozeParts;
  if (!parts?.length) return line?.sentence || "";
  return parts
    .map((part, index) => `${part}${values?.[index] || ""}`)
    .join("")
    .replace(/\s+/g, " ")
    .trim();
}
function getLineAnswerText(line) {
  return fillClozeText(line, line?.answers || []);
}
function getCardAnswerText(card) {
  if (card?.worksheetLines?.length) {
    return card.worksheetLines.map(getLineAnswerText).filter(Boolean).join(" ");
  }
  return card?.en || "";
}
function getRandomEnglishVoice() {
  if (typeof window === "undefined" || !window.speechSynthesis?.getVoices) return null;
  const voices = window.speechSynthesis
    .getVoices()
    .filter((voice) => /^en[-_]/i.test(voice.lang || ""));
  if (!voices.length) return null;
  return voices[Math.floor(Math.random() * voices.length)];
}
function speakText(text, { onStart, onEnd, onError } = {}) {
  const utteranceText = (text || "").trim();
  if (!utteranceText || typeof window === "undefined" || !window.speechSynthesis) {
    onError?.();
    return false;
  }
  window.speechSynthesis.cancel();
  const utterance = new SpeechSynthesisUtterance(utteranceText);
  utterance.lang = "en-US";
  const voice = getRandomEnglishVoice();
  if (voice) {
    utterance.voice = voice;
    utterance.lang = voice.lang || "en-US";
  }
  utterance.rate = 0.88;
  utterance.pitch = 1;
  utterance.onstart = () => onStart?.();
  utterance.onend = () => onEnd?.();
  utterance.onerror = () => onError?.();
  window.speechSynthesis.speak(utterance);
  return true;
}
function playCorrectSound() {
  try {
    const AudioContextClass = window.AudioContext || window.webkitAudioContext;
    if (!AudioContextClass) return;
    const ctx = new AudioContextClass();
    const gain = ctx.createGain();
    gain.gain.setValueAtTime(0.0001, ctx.currentTime);
    gain.gain.exponentialRampToValueAtTime(0.08, ctx.currentTime + 0.02);
    gain.gain.exponentialRampToValueAtTime(0.0001, ctx.currentTime + 0.28);
    gain.connect(ctx.destination);
    [660, 880].forEach((frequency, index) => {
      const oscillator = ctx.createOscillator();
      oscillator.type = "sine";
      oscillator.frequency.setValueAtTime(frequency, ctx.currentTime + index * 0.08);
      oscillator.connect(gain);
      oscillator.start(ctx.currentTime + index * 0.08);
      oscillator.stop(ctx.currentTime + 0.3);
    });
    window.setTimeout(() => ctx.close(), 450);
  } catch {
    // Audio can be blocked by browser policy; answering a question should unlock it, but fail quietly.
  }
}
function normalizeViewerId(value) {
  return (value || "").trim().replace(/^id[:：]/i, "").toLowerCase();
}
function isEnabledCell(value) {
  return getEnabledNumber(value) > 0;
}
function getEnabledNumber(value) {
  const raw = String(value ?? "").trim();
  const numeric = Number(raw);
  return Number.isFinite(numeric) && numeric > 0 ? numeric : 0;
}
function getIdFieldName(viewerId) {
  return "id:" + normalizeViewerId(viewerId);
}
function normalizeIdColumnKey(key) {
  return normKey(key).toLowerCase().replace(/^id：/, "id:");
}
function isIdColumn(key) {
  return normalizeIdColumnKey(key).startsWith("id:");
}
function getViewerIdCellValue(row, viewerId) {
  const normalizedId = normalizeViewerId(viewerId);
  if (!normalizedId || normalizedId === "admin") return "";
  const target = getIdFieldName(normalizedId);
  for (const key of Object.keys(row)) {
    if (isIdColumn(key) && normalizeIdColumnKey(key) === target) return row[key];
  }
  return "";
}
function rowMatchesViewer(row, viewerId) {
  const normalizedId = normalizeViewerId(viewerId);
  if (normalizedId === "admin") return true;
  if (!normalizedId) return false;
  return isEnabledCell(getViewerIdCellValue(row, normalizedId));
}
function resolveLessonDisplayOrders(lessonOrderValues, order) {
  const visibleTitles = order.filter((title) => lessonOrderValues[title]?.length);
  const allOne = visibleTitles.length > 0 && visibleTitles.every((title) => lessonOrderValues[title].every((value) => value === 1));
  if (allOne) {
    const resolved = {};
    visibleTitles.forEach((title, index) => {
      resolved[title] = index + 1;
    });
    return resolved;
  }
  const resolved = {};
  for (const title of visibleTitles) {
    resolved[title] = Math.max(...lessonOrderValues[title]);
  }
  return resolved;
}
function contentKey(viewerId, key) {
  return `content:${normalizeViewerId(viewerId) || "none"}:${key}`;
}
function itemSortValue(item) {
  const raw = (item || "").toString().trim();
  const numeric = Number(raw.match(/\d+(?:\.\d+)?/)?.[0]);
  if (Number.isFinite(numeric)) return numeric;
  const letter = raw.match(/[A-Za-z]/)?.[0]?.toUpperCase();
  return letter ? 100000 + letter.charCodeAt(0) : Number.MAX_SAFE_INTEGER;
}
function addLessonPoint(lessonBucket, item, point, pointImage) {
  const text = (point || "").trim();
  const image = normalizeImageUrl(pointImage);
  if (!text && !image) return;
  const sortValue = itemSortValue(item);
  const key = `${sortValue}:${text}:${image}`;
  if (lessonBucket.pointKeys.has(key)) return;
  lessonBucket.pointKeys.add(key);
  lessonBucket.points.push({ sortValue, text, image });
}
function buildContentFromRows(rows, viewerId = "") {
  const order = [];
  const byTitle = {};
  const visibleTitles = new Set();
  const lessonOrderValues = {};
  const normalizedViewerId = normalizeViewerId(viewerId);
  rows.forEach((row) => {
    const sentence = getField(row, "sentence");
    const en = sentence || getField(row, "en");
    const choices = splitChoices(getField(row, "choices"));
    const answer = getField(row, "answer");
    const cloze = parseClozeSentence(en, answer);
    const audioUrl = normalizeAudioUrl(getField(row, "audio"));
    const point = getField(row, "point");
    const pointImage = getField(row, "pointImage");
    const hasQuestionContent = !!(en || choices.length > 0 || answer || audioUrl);
    if (!hasQuestionContent && !point && !pointImage) return;
    const title = getField(row, "lesson") || "未分類";
    if (rowMatchesViewer(row, normalizedViewerId)) {
      visibleTitles.add(title);
      if (normalizedViewerId !== "admin") {
        const orderValue = getEnabledNumber(getViewerIdCellValue(row, normalizedViewerId));
        if (orderValue > 0) {
          if (!lessonOrderValues[title]) lessonOrderValues[title] = [];
          lessonOrderValues[title].push(orderValue);
        }
      }
    }
    const lessonNo = getField(row, "lessonNo");
    const item = getField(row, "item");
    const emoji = getField(row, "emoji");
    const image = normalizeImageUrl(getField(row, "image"));
    const note = getField(row, "note");
    const hint = getField(row, "hint");
    const hintImage = normalizeImageUrl(getField(row, "hintImage"));
    if (!byTitle[title]) {
      byTitle[title] = { title, lessonNo, emoji: emoji || "📇", cards: [], cardsByItem: {}, points: [], pointKeys: new Set() };
      order.push(title);
    } else if (emoji && byTitle[title].emoji === "📇") {
      byTitle[title].emoji = emoji;
    } else if (lessonNo && !byTitle[title].lessonNo) {
      byTitle[title].lessonNo = lessonNo;
    }
    addLessonPoint(byTitle[title], item, point, pointImage);
    const isWorksheetRow = hasQuestionContent && !!(sentence || choices.length > 0 || answer || audioUrl || item);
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
      if (hint && !card.hint) card.hint = hint;
      if (hintImage && !card.hintImage) card.hintImage = hintImage;
      if (audioUrl && !card.audioUrl) card.audioUrl = audioUrl;
      card.worksheetLines.push({
        sentence: cloze.sentence,
        clozeParts: cloze.clozeParts,
        choices,
        answers: cloze.answers,
        blankCount: cloze.blankCount,
        audioUrl,
        note,
        hint,
        hintImage,
      });
    } else {
      byTitle[title].cards.push({ en, image, emoji, note, hint, hintImage, audioUrl });
    }
  });

  const usedLessonIds = new Set();
  const lessonsMap = {};
  const index = [];
  const resolvedDisplayOrders = normalizedViewerId === "admin" ? {} : resolveLessonDisplayOrders(lessonOrderValues, order);
  order.forEach((title) => {
    if (normalizedViewerId !== "admin" && !visibleTitles.has(title)) return;
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
        hint: c.hint,
        hintImage: c.hintImage,
        visualType: c.image ? "photo" : "emoji",
        photoUrl: c.image || undefined,
        audioUrl: c.audioUrl || undefined,
        worksheetLines: c.worksheetLines || undefined,
        emoji: c.image ? "" : c.emoji || "❓",
      };
    });
    const point = l.points
      .slice()
      .sort((a, b) => a.sortValue - b.sortValue)
      .map((entry) => entry.text)
      .filter(Boolean)
      .join("\n");
    const pointItems = l.points
      .slice()
      .sort((a, b) => a.sortValue - b.sortValue)
      .map(({ text, image }) => ({ text, image }))
      .filter((entry) => entry.text || entry.image);
    const displayOrder = resolvedDisplayOrders[title] || "";
    lessonsMap[lessonId] = { id: lessonId, title, lessonNo: l.lessonNo, displayOrder, emoji: l.emoji, cards, point, pointItems };
    index.push({ id: lessonId, title, lessonNo: l.lessonNo, displayOrder, emoji: l.emoji, count: cards.length, point, pointItems });
  });
  return { index, lessonsMap };
}
function isExcelBuffer(buffer) {
  const bytes = new Uint8Array(buffer.slice(0, 4));
  return bytes[0] === 0x50 && bytes[1] === 0x4b && bytes[2] === 0x03 && bytes[3] === 0x04;
}

function reviewIntervalDays(level) {
  if (level < 90) return 0;
  if (level < 270) return 1;
  if (level < 540) return 2;
  if (level < 1080) return 4;
  if (level < 2160) return 7;
  if (level < 4320) return 14;
  return 30;
}

function getTotalPerfectBadges(perfectByLesson = {}) {
  return Object.values(perfectByLesson).reduce((sum, value) => sum + (Number(value) || 0), 0);
}

function getPersonalLevel(stats = {}) {
  const level = Number(stats.level ?? stats.personalLevel ?? stats.lv);
  if (Number.isFinite(level) && level > 0) return Math.min(MAX_LEVEL, Math.floor(level));
  const xp = Number(stats.xp || 0);
  const badgeCount = getTotalPerfectBadges(stats.perfectByLesson || {});
  return Math.min(MAX_LEVEL, Math.max(1, 1 + Math.floor(xp / XP_PER_PERSONAL_LEVEL) + badgeCount * PERFECT_BADGE_LEVEL_BONUS));
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

function stableStringify(value) {
  if (value === null || typeof value !== "object") return JSON.stringify(value);
  if (Array.isArray(value)) return `[${value.map((entry) => stableStringify(entry)).join(",")}]`;
  return `{${Object.keys(value)
    .sort()
    .map((key) => `${JSON.stringify(key)}:${stableStringify(value[key])}`)
    .join(",")}}`;
}

async function checksumText(text) {
  if (typeof crypto !== "undefined" && crypto.subtle && typeof TextEncoder !== "undefined") {
    const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
    return Array.from(new Uint8Array(digest))
      .map((byte) => byte.toString(16).padStart(2, "0"))
      .join("");
  }
  let hash = 2166136261;
  for (let i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16).padStart(8, "0");
}

async function buildPersonalExport(payload) {
  const checksum = await checksumText(stableStringify(payload));
  return JSON.stringify(
    {
      type: PERSONAL_EXPORT_TYPE,
      version: PERSONAL_EXPORT_VERSION,
      schemaVersion: PERSONAL_EXPORT_VERSION,
      exportedAt: payload.meta?.exportedAt || payload.exportedAt || new Date().toISOString(),
      checksum,
      payload,
    },
    null,
    2
  );
}

function asPlainObject(value) {
  return value && typeof value === "object" && !Array.isArray(value) ? value : {};
}

function normalizeStats(value) {
  const stats = asPlainObject(value);
  const xp = Number(stats.xp ?? stats.totalXp ?? stats.points ?? 0);
  const rawLevel = Number(stats.level ?? stats.personalLevel ?? stats.lv);
  const reviewDates = Array.isArray(stats.reviewDates)
    ? stats.reviewDates
    : Array.isArray(stats.streakDates)
      ? stats.streakDates
      : [];
  return {
    ...stats,
    xp: Number.isFinite(xp) ? xp : 0,
    level: Number.isFinite(rawLevel) && rawLevel > 0 ? Math.min(MAX_LEVEL, Math.floor(rawLevel)) : getPersonalLevel(stats),
    reviewDates,
    perfectByLesson: asPlainObject(stats.perfectByLesson || stats.badges || stats.perfect || {}),
  };
}

function normalizeProgressByLesson(value) {
  const progress = asPlainObject(value);
  const normalized = {};
  for (const [lessonId, lessonProgress] of Object.entries(progress)) {
    if (lessonId && lessonProgress && typeof lessonProgress === "object" && !Array.isArray(lessonProgress)) {
      normalized[lessonId] = lessonProgress;
    }
  }
  return normalized;
}

function normalizePersonalPayload(payload, version = 1) {
  const source = asPlainObject(payload);
  const user = asPlainObject(source.user);
  const learning = asPlainObject(source.learning);
  const prefs = asPlainObject(source.prefs || user.prefs || source.preferences || {});
  return {
    schemaVersion: PERSONAL_EXPORT_VERSION,
    sourceVersion: version,
    exportedAt: source.exportedAt || source.meta?.exportedAt || "",
    viewerId: source.viewerId || user.viewerId || user.id || "",
    prefs: { ...DEFAULT_USER_PREFS, ...prefs },
    stats: normalizeStats(source.stats || learning.stats || source.learningStats || {}),
    progressByLesson: normalizeProgressByLesson(source.progressByLesson || learning.progressByLesson || source.progress || learning.progress || {}),
  };
}

function buildAppInfoText() {
  return [
    `${APP_NAME}`,
    `Version: ${APP_VERSION}`,
    `Generated at: ${new Date().toISOString()}`,
    "",
    "Copyright:",
    "Copyright (c) 2026 Kazuro Ueshima.",
    "",
    "Application License:",
    "BSD 3-Clause License",
    "",
    "Redistribution and use in source and binary forms, with or without modification, are permitted provided that the following conditions are met:",
    "1. Redistributions of source code must retain the above copyright notice, this list of conditions and the following disclaimer.",
    "2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions and the following disclaimer in the documentation and/or other materials provided with the distribution.",
    "3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or promote products derived from this software without specific prior written permission.",
    "",
    "THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS \"AS IS\" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED.",
    "",
    "Third-party software:",
    "This application uses third-party packages subject to their respective licenses.",
    "- React 19.2.7: MIT License",
    "- lucide-react 1.23.0: ISC License",
    "- PapaParse 5.5.4: MIT License",
    "- SheetJS xlsx 0.18.5: Apache-2.0 License",
    "- Vite and build tooling: subject to their respective package licenses",
    "",
    "Notes:",
    "Lesson content, worksheet images, audio files, and imported spreadsheet data may have separate rights and permissions.",
  ].join("\n");
}

function downloadTextFile(filename, text) {
  const blob = new Blob([text], { type: "text/plain;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = filename;
  document.body.appendChild(link);
  link.click();
  link.remove();
  URL.revokeObjectURL(url);
}

async function parsePersonalExport(text) {
  const data = JSON.parse(text || "{}");
  const version = Number(data.version);
  if (data.type !== PERSONAL_EXPORT_TYPE || !Number.isFinite(version) || version < 1 || version > PERSONAL_EXPORT_VERSION || !data.payload || !data.checksum) {
    throw new Error("このアプリの個人設定データではありません");
  }
  const expected = await checksumText(stableStringify(data.payload));
  if (expected !== data.checksum) {
    throw new Error("チェックサムが一致しません。内容が変更されている可能性があります");
  }
  return normalizePersonalPayload(data.payload, version);
}

export default function App() {
  const [ready, setReady] = useState(false);
  const [lessons, setLessons] = useState({});
  const [index, setIndex] = useState([]);
  const [stats, setStats] = useState({ xp: 0, reviewDates: [] });
  const [progressByLesson, setProgressByLesson] = useState({});
  const [screen, setScreen] = useState({ name: "home" });
  const [viewerId, setViewerId] = useState("");
  const [userPrefs, setUserPrefs] = useState(DEFAULT_USER_PREFS);

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

  const loadContentForViewer = async (activeViewerId) => {
    const normalizedId = normalizeViewerId(activeViewerId);
    const idxRaw = await safeGet(contentKey(normalizedId, "lesson-index"), false);
    let idx = idxRaw ? JSON.parse(idxRaw) : [];

    let loaded = {};
    for (const meta of idx) {
      const raw = await safeGet(contentKey(normalizedId, "lesson:" + meta.id), false);
      if (raw) loaded[meta.id] = JSON.parse(raw);
    }

    try {
      const response = await fetch(DEFAULT_EXCEL_URL, { cache: "no-cache" });
      if (!response.ok) throw new Error("Default Excel not found");
      const rows = parseSheetBuffer(
        await response.arrayBuffer(),
        "review.xlsx",
        response.headers.get("content-type") || "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      const { index: defaultIndex, lessonsMap } = buildContentFromRows(rows, normalizedId);
      for (const oldMeta of idx) {
        if (!lessonsMap[oldMeta.id]) await safeDelete(contentKey(normalizedId, "lesson:" + oldMeta.id), false);
      }
      for (const l of Object.values(lessonsMap)) {
        await safeSet(contentKey(normalizedId, "lesson:" + l.id), JSON.stringify(l), false);
      }
      await safeSet(contentKey(normalizedId, "lesson-index"), JSON.stringify(defaultIndex), false);
      await safeSet(contentKey(normalizedId, CONTENT_IMPORTED_KEY), CONTENT_VERSION, false);
      idx = defaultIndex;
      loaded = lessonsMap;
    } catch {
      // Keep the app usable with the last successful fetch.
    }

    setIndex(idx);
    setLessons(loaded);
    setProgressByLesson({});
    return idx;
  };

  useEffect(() => {
    if (typeof window !== "undefined" && window.speechSynthesis?.getVoices) {
      window.speechSynthesis.getVoices();
      window.speechSynthesis.onvoiceschanged = () => window.speechSynthesis.getVoices();
    }
    (async () => {
      const statsRaw = await safeGet("stats", false);
      const st = statsRaw ? JSON.parse(statsRaw) : { xp: 0, reviewDates: [] };
      const prefsRaw = await safeGet(USER_PREFS_KEY, false);
      const prefs = prefsRaw ? { ...DEFAULT_USER_PREFS, ...JSON.parse(prefsRaw) } : DEFAULT_USER_PREFS;
      const storedViewerId = normalizeViewerId(await safeGet(USER_ID_KEY, false));
      setViewerId(storedViewerId);
      setStats(st);
      setUserPrefs(prefs);
      if (storedViewerId) await loadContentForViewer(storedViewerId);
      setReady(true);
    })();
    return () => {
      if (typeof window !== "undefined" && window.speechSynthesis) window.speechSynthesis.onvoiceschanged = null;
    };
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

  const persistUserPrefs = useCallback(async (prefs) => {
    const nextPrefs = { ...DEFAULT_USER_PREFS, ...prefs };
    setUserPrefs(nextPrefs);
    await safeSet(USER_PREFS_KEY, JSON.stringify(nextPrefs), false);
  }, []);

  const resetPersonalSettings = async (lessonIndex = []) => {
    const resetStats = { xp: 0, level: 1, reviewDates: [], perfectByLesson: {} };
    const lessonIds = new Set([
      ...index.map((meta) => meta.id),
      ...lessonIndex.map((meta) => meta.id),
      ...Object.keys(progressByLesson),
    ]);
    await safeDelete("stats", false);
    await safeDelete(USER_PREFS_KEY, false);
    for (const lessonId of lessonIds) {
      await safeDelete("progress:" + lessonId, false);
    }
    setStats(resetStats);
    setUserPrefs(DEFAULT_USER_PREFS);
    setProgressByLesson({});
  };

  // ---- spreadsheet sync ----
  const importRows = async (rows) => {
    const activeViewerId = normalizeViewerId(viewerId);
    if (!activeViewerId) throw new Error("IDを入力してください");
    const { index: newIndex, lessonsMap: newLessons } = buildContentFromRows(rows, activeViewerId);
    for (const oldMeta of index) {
      if (!newLessons[oldMeta.id]) await safeDelete(contentKey(activeViewerId, "lesson:" + oldMeta.id), false);
    }
    for (const l of Object.values(newLessons)) {
      await safeSet(contentKey(activeViewerId, "lesson:" + l.id), JSON.stringify(l), false);
    }
    await safeSet(contentKey(activeViewerId, "lesson-index"), JSON.stringify(newIndex), false);
    await safeSet(contentKey(activeViewerId, CONTENT_IMPORTED_KEY), CONTENT_VERSION, false);
    setIndex(newIndex);
    setLessons(newLessons);
    return newIndex;
  };

  const saveViewerId = async (rawId) => {
    const normalizedId = normalizeViewerId(rawId);
    if (!normalizedId) return;
    await safeSet(USER_ID_KEY, normalizedId, false);
    setViewerId(normalizedId);
    setScreen({ name: "home" });
    setReady(false);
    const loadedIndex = await loadContentForViewer(normalizedId);
    if (normalizedId === "admin") await resetPersonalSettings(loadedIndex);
    setReady(true);
  };

  const exportPersonalSettings = async () => {
    const allProgress = {};
    for (const meta of index) {
      const lessonProgress = progressByLesson[meta.id] || JSON.parse((await safeGet("progress:" + meta.id, false)) || "{}");
      if (Object.keys(lessonProgress).length > 0) allProgress[meta.id] = lessonProgress;
    }
    return buildPersonalExport({
      meta: {
        schemaVersion: PERSONAL_EXPORT_VERSION,
        exportedAt: new Date().toISOString(),
      },
      user: {
        viewerId,
        prefs: userPrefs,
      },
      learning: {
        stats,
        progressByLesson: allProgress,
      },
    });
  };

  const importPersonalSettings = async (text) => {
    const payload = await parsePersonalExport(text);
    const nextViewerId = normalizeViewerId(payload.viewerId || viewerId);
    const nextPrefs = { ...DEFAULT_USER_PREFS, ...(payload.prefs || {}) };
    const nextStats = {
      xp: Number(payload.stats?.xp || 0),
      level: getPersonalLevel(payload.stats || {}),
      reviewDates: Array.isArray(payload.stats?.reviewDates) ? payload.stats.reviewDates : [],
      perfectByLesson: payload.stats?.perfectByLesson || {},
    };
    const nextProgress = payload.progressByLesson && typeof payload.progressByLesson === "object" ? payload.progressByLesson : {};

    if (nextViewerId) {
      await safeSet(USER_ID_KEY, nextViewerId, false);
      setViewerId(nextViewerId);
      await loadContentForViewer(nextViewerId);
    }
    await persistUserPrefs(nextPrefs);
    await persistStats(nextStats);
    for (const [lessonId, progress] of Object.entries(nextProgress)) {
      await safeSet("progress:" + lessonId, JSON.stringify(progress), false);
    }
    setProgressByLesson(nextProgress);
    setScreen({ name: "home" });
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
    <div style={{ fontFamily: "'Inter', sans-serif" }} className={`min-h-screen w-full tone-${userPrefs.tone || "fresh"}`}>
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
        .tone-soft .cabinet-bg {
          background-color: #fffaf7;
          background-image:
            radial-gradient(circle at 16% 12%, rgba(255,184,108,0.16), transparent 26%),
            linear-gradient(180deg, #ffffff 0%, #fff3ec 58%, #fffaf7 100%);
        }
        .tone-soft .drawer-front { background: linear-gradient(180deg, #ffffff 0%, #fff1e9 100%); border-color: #f0c7ad; }
        .tone-soft .brass { background: linear-gradient(180deg, #fff0c8 0%, #ffd489 100%); border-color: #e7b55d; }
        .tone-night .cabinet-bg {
          background-color: #122033;
          background-image:
            radial-gradient(circle at 20% 10%, rgba(91,188,214,0.2), transparent 26%),
            linear-gradient(180deg, #17263b 0%, #101b2b 100%);
        }
        .tone-night .card-paper { background: rgba(255,255,255,0.95); }
        .tone-night .drawer-front { background: linear-gradient(180deg, #f8fbff 0%, #dcecff 100%); border-color: #91b7db; }
        .tone-night .brass { background: linear-gradient(180deg, #d9f6ff 0%, #8dd6ed 100%); border-color: #62bad5; }
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
      ) : !viewerId ? (
        <IdGate onSubmit={saveViewerId} />
      ) : (
        <div className="cabinet-bg min-h-screen">
          <TopBar screen={screen} setScreen={setScreen} xp={stats.xp} level={getPersonalLevel(stats)} streak={streak} />
          <div className="max-w-2xl mx-auto px-4 pb-16">
            {screen.name === "home" && (
              <Home
                index={index}
                perfectByLesson={stats.perfectByLesson || {}}
                onOpen={(id) => setScreen({ name: "review", id })}
              />
            )}
            {screen.name === "settings" && (
              <SettingsScreen
                viewerId={viewerId}
                onChangeViewerId={saveViewerId}
                tone={userPrefs.tone || "fresh"}
                onChangeTone={(tone) => persistUserPrefs({ ...userPrefs, tone })}
                onImportRows={importRows}
                onExportPersonalSettings={exportPersonalSettings}
                onImportPersonalSettings={importPersonalSettings}
              />
            )}
            {screen.name === "lesson" && lessons[screen.id] && (
              <LessonDetail
                lesson={lessons[screen.id]}
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
                  const perfectByLesson = { ...(stats.perfectByLesson || {}) };
                  let xpEarned = result.xpEarned;
                  const currentLevel = getPersonalLevel(stats);
                  const isLowScore = result.total > 0 && result.correct / result.total < 0.5;
                  if (result.total > 0 && result.correct === result.total) {
                    const previousPerfectCount = perfectByLesson[screen.id] || 0;
                    perfectByLesson[screen.id] = previousPerfectCount + 1;
                    xpEarned += previousPerfectCount >= 3 ? EXTRA_PERFECT_BADGE_XP_BONUS : PERFECT_BADGE_XP_BONUS;
                  } else if (isLowScore && (perfectByLesson[screen.id] || 0) > 0) {
                    perfectByLesson[screen.id] -= 1;
                    if (perfectByLesson[screen.id] < 1) delete perfectByLesson[screen.id];
                  }
                  const totalXp = (Number(stats.xp) || 0) + xpEarned;
                  const levelGained = Math.floor(totalXp / XP_LEVEL_THRESHOLD);
                  const nextXp = totalXp % XP_LEVEL_THRESHOLD;
                  const levelLost = isLowScore ? 1 : 0;
                  const nextLevel = Math.max(1, Math.min(MAX_LEVEL, currentLevel + levelGained - levelLost));
                  const nextResult = { ...result, xpEarned, levelGained, levelLost, levelDelta: nextLevel - currentLevel };
                  await persistStats({ ...stats, xp: nextXp, level: nextLevel, reviewDates: Array.from(dates), perfectByLesson });
                  setScreen({ name: "summary", id: screen.id, result: nextResult });
                }}
                onExit={() => setScreen({ name: "home" })}
              />
            )}
            {screen.name === "summary" && (
              <Summary result={screen.result} onReviewAgain={() => setScreen({ name: "review", id: screen.id })} onHome={() => setScreen({ name: "home" })} />
            )}
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------- TopBar ---------------- */
function TopBar({ screen, setScreen, xp, level, streak }) {
  return (
    <div className="sticky top-0 z-30 backdrop-blur-sm bg-white/90 border-b border-[#d7eef6]">
      <div className="max-w-2xl mx-auto px-4 py-3 flex items-center justify-between gap-2">
        <button onClick={() => setScreen({ name: "home" })} className="flex items-center gap-2 text-[#16475f] min-w-0">
          <span className="font-display text-lg font-bold tracking-tight truncate">GDM 復習カード</span>
        </button>
        <div className="flex items-center gap-3 font-mono text-sm shrink-0">
          <div className="flex items-center gap-1 text-[#1687a7]"><Flame size={16} /> {streak}</div>
          <div className="flex items-center gap-1 text-[#16805d]"><Star size={16} /> {xp}</div>
          <div className="rounded bg-[#d7f5eb] px-2 py-1 text-[11px] font-bold text-[#16805d]">Lv {level}</div>
          <button onClick={() => setScreen({ name: "settings" })} className="flex items-center gap-1 text-[10px] bg-[#1687a7] text-white px-2 py-1 rounded" title="設定">
            <Settings size={12} /> 設定
          </button>
          {screen.name !== "home" && (
            <button
              onClick={() => setScreen({ name: "home" })}
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

function IdGate({ onSubmit }) {
  const [value, setValue] = useState("");
  return (
    <div className="cabinet-bg min-h-screen flex items-center justify-center px-4">
      <form
        onSubmit={(event) => {
          event.preventDefault();
          onSubmit(value);
        }}
        className="card-paper w-full max-w-sm rounded-md border border-[#b7d6e6] p-5 shadow-md"
      >
        <div className="mb-4">
          <div className="font-display text-xl font-bold text-[#16475f]">ID</div>
        </div>
        <input
          value={value}
          onChange={(event) => setValue(event.target.value)}
          autoFocus
          autoCapitalize="none"
          autoCorrect="off"
          className="mb-3 w-full rounded border border-[#b7d6e6] bg-white px-3 py-2 text-[#16475f] outline-none focus:border-[#1687a7]"
        />
        <button type="submit" className="w-full rounded bg-[#1687a7] px-3 py-2 font-bold text-white hover:brightness-110">
          はじめる
        </button>
      </form>
    </div>
  );
}

/* ---------------- Home ---------------- */
function Home({ index, perfectByLesson, onOpen }) {
  const [jumpNo, setJumpNo] = useState("");
  const [pointLesson, setPointLesson] = useState(null);
  const displayIndex = useMemo(
    () =>
      index
        .map((meta, originalIndex) => ({
          ...meta,
          displayNo: meta.displayOrder || meta.lessonNo || originalIndex + 1,
        }))
        .sort((a, b) => Number(b.displayNo) - Number(a.displayNo)),
    [index]
  );
  const openByNumber = () => {
    const normalized = jumpNo.trim();
    if (!normalized) return;
    const found = index.find((meta, indexPosition) => String(meta.lessonNo || indexPosition + 1) === normalized);
    if (found && (found.count || 0) > 0) onOpen(found.id);
  };

  return (
    <div className="pt-6">
      <p className="text-[11px] text-[#42677a] mb-5 font-mono">復習の記録(習熟度・XP)は自分だけに保存されます。</p>

      {index.length > 0 && (
        <div className="mb-5 flex items-center gap-2 rounded-md border border-[#b7d6e6] bg-white/80 p-2 shadow-sm">
          <input
            type="number"
            inputMode="numeric"
            min="1"
            value={jumpNo}
            onChange={(event) => setJumpNo(event.target.value)}
            onKeyDown={(event) => {
              if (event.key === "Enter") openByNumber();
            }}
            placeholder="番号"
            className="min-w-0 flex-1 rounded border border-[#b7d6e6] bg-white px-3 py-2 text-sm text-[#16475f] outline-none focus:border-[#1687a7]"
          />
          <button
            type="button"
            onClick={openByNumber}
            className="shrink-0 rounded bg-[#1687a7] px-3 py-2 text-sm font-bold text-white hover:brightness-110"
          >
            ジャンプ
          </button>
        </div>
      )}

      {index.length === 0 && (
        <div className="card-paper rounded-md p-6 text-center mb-4 border border-[#b7d6e6]">
          <p className="font-display text-lg text-[#16475f] mb-1">まだレッスンがありません</p>
          <p className="text-sm text-[#42677a]">右上の設定からCSV/Excelを読み込んでください</p>
        </div>
      )}

      <div className="space-y-3">
        {displayIndex.map((meta) => (
          <div key={meta.id} className="drawer-front w-full rounded-md p-4 shadow-md">
            <div className="flex items-center gap-4">
              <button onClick={() => onOpen(meta.id)} disabled={(meta.count || 0) < 1} className="flex min-w-0 flex-1 items-center gap-4 text-left hover:brightness-110 transition disabled:opacity-50 disabled:cursor-not-allowed">
                <div className="brass rounded w-12 h-12 flex flex-col items-center justify-center shrink-0">
                  <span className="font-mono text-[11px] font-bold leading-none text-[#16475f]">{meta.displayNo}</span>
                  <span className="text-lg leading-none">{meta.emoji || "📇"}</span>
                </div>
                <div className="flex-1 min-w-0">
                  <div className="font-display text-[#16475f] text-lg font-semibold truncate">{meta.title}</div>
                  <div className="font-mono text-xs text-[#42677a]">{meta.count || 0} 枚のカード</div>
                  <PerfectBadges count={perfectByLesson[meta.id] || 0} />
                </div>
                <ChevronRight className="text-[#1687a7]" size={20} />
              </button>
              {(meta.point || meta.pointItems?.length > 0) && (
                <button
                  type="button"
                  onClick={() => setPointLesson(meta)}
                  className="shrink-0 rounded border border-[#73bfd7] bg-white px-3 py-2 text-xs font-bold text-[#166078] hover:bg-[#e8f7fb]"
                >
                  ポイント
                </button>
              )}
            </div>
          </div>
        ))}
      </div>
      {pointLesson && <PointModal lesson={pointLesson} onClose={() => setPointLesson(null)} />}
    </div>
  );
}

function PointModal({ lesson, onClose }) {
  const pointItems = lesson.pointItems?.length
    ? lesson.pointItems
    : (lesson.point || "")
      .split("\n")
      .filter(Boolean)
      .map((text) => ({ text, image: "" }));

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-[#16475f]/30 px-4">
      <div className="card-paper max-h-[86vh] w-full max-w-md overflow-y-auto rounded-md border border-[#b7d6e6] p-5 shadow-xl">
        <div className="mb-3 flex items-start justify-between gap-3">
          <div>
            <div className="font-display text-lg font-bold text-[#16475f]">{lesson.title}</div>
            <div className="font-mono text-xs text-[#42677a]">ポイント</div>
          </div>
          <button type="button" onClick={onClose} className="rounded px-2 py-1 text-[#42677a] hover:bg-[#e8f7fb]">
            閉じる
          </button>
        </div>
        <div className="space-y-4 text-sm leading-7 text-[#16475f]">
          {pointItems.map((point, index) => (
            <div key={`${point.text}-${point.image}-${index}`} className="space-y-2">
              {point.image && (
                <img
                  src={point.image}
                  alt={point.text || "ポイント画像"}
                  className="max-h-64 w-full rounded border border-[#b7d6e6] bg-white object-contain"
                />
              )}
              {point.text && <div className="whitespace-pre-line">{point.text}</div>}
            </div>
          ))}
        </div>
      </div>
    </div>
  );
}

function PerfectBadges({ count }) {
  if (!count) return null;
  const visible = Math.min(count, 5);
  const medalFor = (index) => (index === 0 ? "🥉" : index === 1 ? "🥈" : "🥇");
  const labelFor = (index) => (index === 0 ? "銅メダル" : index === 1 ? "銀メダル" : "金メダル");
  return (
    <div className="mt-2 flex items-center gap-1.5" aria-label={`全問正解 ${count}回`}>
      {Array.from({ length: visible }).map((_, index) => (
        <span
          key={index}
          title={labelFor(index)}
          className="inline-flex h-7 w-7 items-center justify-center rounded-full border border-white/80 bg-white text-[17px] shadow-md ring-1 ring-[#b7d6e6]"
        >
          {medalFor(index)}
        </span>
      ))}
      {count > visible && <span className="rounded-full bg-[#d7f5eb] px-2 py-0.5 font-mono text-[10px] font-bold text-[#16805d]">+{count - visible}</span>}
    </div>
  );
}

function SettingsScreen({
  viewerId,
  onChangeViewerId,
  tone,
  onChangeTone,
  onImportRows,
  onExportPersonalSettings,
  onImportPersonalSettings,
}) {
  const [showTeacherSettings, setShowTeacherSettings] = useState(false);
  return (
    <div className="space-y-4 pt-6">
      <ViewerIdPanel viewerId={viewerId} onChangeViewerId={onChangeViewerId} />
      <TonePanel tone={tone} onChangeTone={onChangeTone} />
      <PersonalSettingsPanel onExport={onExportPersonalSettings} onImport={onImportPersonalSettings} />
      <AppInfoPanel />
      <button
        type="button"
        onClick={() => setShowTeacherSettings((value) => !value)}
        className="w-full rounded-md border border-[#73bfd7] bg-white/90 px-4 py-3 text-left font-display font-bold text-[#166078] shadow-sm hover:bg-[#e8f7fb]"
      >
        {showTeacherSettings ? "設定(先生用)を閉じる" : "設定(先生用)"}
      </button>
      {showTeacherSettings && <SheetSyncPanel onImportRows={onImportRows} />}
    </div>
  );
}

function AppInfoPanel() {
  return (
    <div className="card-paper rounded-md border border-[#b7d6e6] p-4">
      <div className="mb-2 font-display font-bold text-[#16475f]">アプリ情報</div>
      <p className="mb-3 text-xs leading-5 text-[#42677a]">
        ソフトバージョン、Copyright、ライセンス、利用している主なライブラリの情報をテキストでダウンロードできます。
      </p>
      <button
        type="button"
        onClick={() => downloadTextFile("gdm-review-app-info.txt", buildAppInfoText())}
        className="flex w-full items-center justify-center gap-2 rounded border border-[#73bfd7] bg-white px-3 py-2 text-sm font-bold text-[#166078] hover:bg-[#e8f7fb]"
      >
        <Download size={16} />
        アプリ情報をダウンロード
      </button>
    </div>
  );
}

function ViewerIdPanel({ viewerId, onChangeViewerId }) {
  const [value, setValue] = useState(viewerId || "");
  return (
    <form
      onSubmit={(event) => {
        event.preventDefault();
        onChangeViewerId(value);
      }}
      className="card-paper rounded-md border border-[#b7d6e6] p-4"
    >
      <div className="mb-2 font-display font-bold text-[#16475f]">ID</div>
      <div className="flex gap-2">
        <input
          value={value}
          onChange={(event) => setValue(event.target.value)}
          autoCapitalize="none"
          autoCorrect="off"
          className="min-w-0 flex-1 rounded border border-[#b7d6e6] bg-white px-3 py-2 text-sm text-[#16475f] outline-none focus:border-[#1687a7]"
        />
        <button type="submit" className="shrink-0 rounded bg-[#1687a7] px-3 py-2 text-sm font-bold text-white hover:brightness-110">
          変更
        </button>
      </div>
    </form>
  );
}

function TonePanel({ tone, onChangeTone }) {
  const tones = [
    { id: "fresh", label: "さわやか", swatch: "bg-[#73bfd7]" },
    { id: "soft", label: "やわらか", swatch: "bg-[#ffd489]" },
    { id: "night", label: "夜", swatch: "bg-[#17263b]" },
  ];
  return (
    <div className="card-paper rounded-md border border-[#b7d6e6] p-4">
      <div className="mb-3 flex items-center gap-2 font-display font-bold text-[#16475f]">
        <Palette size={18} />
        画面トーン
      </div>
      <div className="grid grid-cols-3 gap-2">
        {tones.map((option) => {
          const selected = tone === option.id;
          return (
            <button
              key={option.id}
              type="button"
              onClick={() => onChangeTone(option.id)}
              className={
                "flex items-center justify-center gap-2 rounded border px-3 py-2 text-sm font-bold " +
                (selected ? "border-[#1687a7] bg-[#e8f7fb] text-[#16475f]" : "border-[#b7d6e6] bg-white text-[#42677a] hover:bg-[#f4fbfd]")
              }
            >
              <span className={`h-3 w-3 rounded-full ${option.swatch}`} />
              {option.label}
            </button>
          );
        })}
      </div>
    </div>
  );
}

function PersonalSettingsPanel({ onExport, onImport }) {
  const [status, setStatus] = useState(null);

  const handleExport = async () => {
    setStatus({ type: "loading", message: "エクスポート中..." });
    try {
      const exported = await onExport();
      downloadTextFile("gdm-review-personal-settings.txt", exported);
      setStatus({ type: "success", message: "個人設定をエクスポートしました" });
    } catch (error) {
      setStatus({ type: "error", message: "エクスポートに失敗しました: " + (error?.message || "不明なエラー") });
    }
  };

  const handleImportFile = async (file) => {
    if (!file) return;
    setStatus({ type: "loading", message: "インポート中..." });
    try {
      await onImport(await file.text());
      setStatus({ type: "success", message: "個人設定をインポートしました" });
    } catch (error) {
      setStatus({ type: "error", message: "インポートに失敗しました: " + (error?.message || "不明なエラー") });
    }
  };

  return (
    <div className="card-paper rounded-md border border-[#b7d6e6] p-4">
      <div className="mb-2 font-display font-bold text-[#16475f]">個人設定の引き継ぎ</div>
      <p className="mb-3 text-xs leading-5 text-[#42677a]">
        XP、Lv、復習状況、バッチ、画面トーンをテキストファイルとして保存できます。機種変更時はエクスポートしたファイルを新しい端末でインポートしてください。
      </p>
      <div className="mb-3 rounded border border-[#b7d6e6] bg-white/70 p-3 text-xs leading-5 text-[#42677a]">
        <div className="mb-1 font-display font-bold text-[#16475f]">表示の意味</div>
        <div>炎: 連続して復習した日数です。</div>
        <div>星: XPです。正解するほど増える学習ポイントです。</div>
        <div>メダル: レッスンを全問正解した回数です。正答率が5割未満だと1つ減ることがあります。</div>
        <div>Lv: 個人レベルです。復習状況に応じて上がり、正答率が5割未満だと下がることがあります。</div>
      </div>
      <div className="grid grid-cols-2 gap-2">
        <button
          type="button"
          onClick={handleExport}
          className="flex items-center justify-center gap-2 rounded bg-[#1687a7] px-3 py-2 text-sm font-bold text-white hover:brightness-110"
        >
          <Download size={16} />
          エクスポート
        </button>
        <label className="flex cursor-pointer items-center justify-center gap-2 rounded border border-[#73bfd7] bg-white px-3 py-2 text-sm font-bold text-[#166078] hover:bg-[#e8f7fb]">
          <Upload size={16} />
          インポート
          <input
            type="file"
            accept=".txt,.json,text/plain,application/json"
            className="hidden"
            onChange={(event) => {
              handleImportFile(event.target.files?.[0]);
              event.target.value = "";
            }}
          />
        </label>
      </div>
      {status && (
        <p className={"mt-2 text-xs " + (status.type === "error" ? "text-[#b42335]" : "text-[#16805d]")}>{status.message}</p>
      )}
    </div>
  );
}

/* ---------------- Settings / Import Panel ---------------- */
function SheetSyncPanel({ onImportRows }) {
  const [status, setStatus] = useState(null);
  const [showHelp, setShowHelp] = useState(false);
  const csvExample = `lesson_no,lesson,item,image,sentence,choices,answer,audio,note,hint,hint_image,point,point_image,ID:sample
1,WorkP22,1,https://example.com/picture1.png,This is a {{man}}.,man|woman,,https://example.com/audio1.mp3,{{ }} の中が空欄になります,人物を表す語を選びます,https://example.com/hint1.png,be動詞の後ろは名詞を置けます,https://example.com/point1.png,1
1,WorkP22,1,https://example.com/picture1.png,Her {{book}} is on the {{table}}.,book|bag|table|desk,,,,,,,,1
1,WorkP22,2,https://example.com/picture2.png,This is a ____.,glass|bottle,glass,,____ を使う場合は answer に正解を書きます,,,,,1`;

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
        <span className="font-mono"> lesson_no / lesson / item / image / sentence / choices / answer / audio / note / hint / hint_image / point / point_image / ID:任意のID </span>
        を使えます。image、hint_image、point_image、audioにはGoogle Driveに配置した画像と音声をURLで指定してください。Google Driveの上位フォルダ設定はリンクを知っている人が閲覧できる設定にしてください。
      </p>
      <p className="mb-3 text-xs leading-5 text-[#42677a]">
        audio列は、ブラウザの読み上げではなく音声をカスタマイズしたい場合だけリンクを書いてください。未入力の場合は正解の英文を読み上げます。
      </p>
      <p className="text-xs text-[#42677a] mb-3">
        ID列は <span className="font-mono">ID:sample</span> のように作ります。入力IDと同じ列に数値があるレッスンだけ表示します。
        その数値が1,2,3...なら数値の降順で並びます。すべて1の場合は、シートの上から1,2,3...として扱い、降順に表示します。
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
          <span className="font-mono"> man|woman|girlまたはman/woman/girl </span>
          のように区切って書きます。<span className="font-mono">hint_image</span> と <span className="font-mono">point_image</span> にURLを書くと、ヒントやレッスン一覧の「ポイント」内に画像も表示されます。
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

function AudioButton({ src, text, label = "音声" }) {
  const audioRef = useRef(null);
  const [isPlaying, setIsPlaying] = useState(false);
  const [hasError, setHasError] = useState(false);
  const drivePreviewUrl = getGoogleDrivePreviewUrl(src);
  const playableSrc = normalizeAudioUrl(src);
  const speechText = (text || "").trim();

  useEffect(() => {
    return () => {
      if (audioRef.current) {
        audioRef.current.pause();
        audioRef.current = null;
      }
      if (typeof window !== "undefined" && window.speechSynthesis) window.speechSynthesis.cancel();
    };
  }, []);

  if (!playableSrc && !speechText) return null;

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
    if (!playableSrc && speechText) {
      if (isPlaying && window.speechSynthesis) {
        window.speechSynthesis.cancel();
        setIsPlaying(false);
        return;
      }
      const ok = speakText(speechText, {
        onStart: () => setIsPlaying(true),
        onEnd: () => setIsPlaying(false),
        onError: () => {
          setIsPlaying(false);
          setHasError(true);
        },
      });
      if (!ok) setHasError(true);
      return;
    }

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
          {(line.audioUrl || getLineAnswerText(line)) && <div className="mt-1"><AudioButton src={line.audioUrl} text={getLineAnswerText(line)} /></div>}
        </div>
      ))}
    </div>
  );
}

function HintContent({ text, image }) {
  return (
    <div className="space-y-2">
      {image && (
        <img
          src={image}
          alt={text || "ヒント画像"}
          className="max-h-56 w-full rounded border border-[#b7d6e6] bg-white object-contain"
        />
      )}
      {text && <div className="whitespace-pre-line">{text}</div>}
    </div>
  );
}

function WorksheetQuestion({ card, selected, onSubmit }) {
  const [fills, setFills] = useState({});
  const [visibleHints, setVisibleHints] = useState({});
  const lines = card.worksheetLines || [];

  useEffect(() => {
    setFills({});
    setVisibleHints({});
  }, [card.id]);

  const shuffledChoicesByLine = useMemo(
    () => lines.map((line) => shuffleList(line.choices || [])),
    [card.id, lines]
  );
  const getBlankCount = (line) => line.blankCount ?? (line.answers || []).length ?? Math.max((line.clozeParts || []).length - 1, 0);
  const blankIndexes = (line) => Array.from({ length: getBlankCount(line) }, (_, index) => index);
  const allFilled = lines.every((line, lineIndex) =>
    blankIndexes(line).every((blankIndex) => fills[`${lineIndex}-${blankIndex}`])
  );

  const setBlank = (lineIndex, blankIndex, choice) => {
    if (selected) return;
    setFills((prev) => ({ ...prev, [`${lineIndex}-${blankIndex}`]: choice }));
  };

  const isCorrect = () =>
    lines.every((line, lineIndex) =>
      blankIndexes(line).every((blankIndex) => isAnswerCorrect((line.answers || [])[blankIndex], fills[`${lineIndex}-${blankIndex}`]))
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
                  {partIndex < getBlankCount(line) && (
                    <span className={`mx-1 inline-flex min-w-20 items-center justify-center border-b-2 px-2 ${selected ? isAnswerCorrect((line.answers || [])[partIndex], fills[`${lineIndex}-${partIndex}`]) ? "border-[#16805d] text-[#16805d]" : "border-[#b42335] text-[#b42335]" : "border-[#1687a7] text-[#16475f]"}`}>
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
                    const nextBlank = blankIndexes(line).find((blankIndex) => !fills[`${lineIndex}-${blankIndex}`]);
                    if (nextBlank >= 0) setBlank(lineIndex, nextBlank, choice);
                  }}
                  className="rounded border border-[#b7d6e6] bg-white px-2 py-1 text-xs text-[#166078] hover:bg-[#e8f7fb] disabled:opacity-70"
                >
                  {choice}
                </button>
              ))}
              {blankIndexes(line).some((blankIndex) => fills[`${lineIndex}-${blankIndex}`]) && !selected && (
                <button
                  type="button"
                  onClick={() => setFills((prev) => {
                    const next = { ...prev };
                    blankIndexes(line).forEach((blankIndex) => delete next[`${lineIndex}-${blankIndex}`]);
                    return next;
                  })}
                  className="rounded px-2 py-1 text-xs text-[#42677a] underline"
                >
                  クリア
                </button>
              )}
            </div>
            {(line.hint || line.hintImage) && (
              <div className="mt-2 text-left">
                <button
                  type="button"
                  onClick={() => setVisibleHints((prev) => ({ ...prev, [lineIndex]: !prev[lineIndex] }))}
                  className="rounded border border-[#73bfd7] bg-white px-2 py-1 text-xs font-bold text-[#166078] hover:bg-[#e8f7fb]"
                >
                  ヒント
                </button>
                {visibleHints[lineIndex] && (
                  <div className="mt-2 rounded border border-[#b7d6e6] bg-white/80 p-2 text-xs leading-5 text-[#42677a]">
                    <HintContent text={line.hint} image={line.hintImage} />
                  </div>
                )}
              </div>
            )}
            {(line.audioUrl || getLineAnswerText(line)) && <div className="mt-2"><AudioButton src={line.audioUrl} text={getLineAnswerText(line)} /></div>}
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
function LessonDetail({ lesson, onEnsureProgress, onBack, onStartReview }) {
  useEffect(() => {
    onEnsureProgress();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [lesson.id]);

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
                  {(c.audioUrl || c.en) && <div className="mt-1"><AudioButton src={c.audioUrl} text={c.en} /></div>}
                </>
              )}
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}

function MasteryDots({ level }) {
  const pct = Math.min(100, Math.max(0, (level / MAX_LEVEL) * 100));
  return (
    <div className="hidden min-w-[76px] items-center gap-2 sm:flex" aria-label={`レベル ${level}`}>
      <span className="font-mono text-[10px] font-bold text-[#16805d]">Lv {level}</span>
      <span className="h-1.5 w-8 overflow-hidden rounded-full bg-[#d7eef6]">
        <span className="block h-full rounded-full bg-[#16805d]" style={{ width: `${pct}%` }} />
      </span>
    </div>
  );
}

/* ---------------- Review Session ---------------- */
function buildQueue(lesson) {
  return [...lesson.cards];
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
  const queue = useRef(buildQueue(lesson));
  const [qIndex, setQIndex] = useState(0);
  const [question, setQuestion] = useState(() => makeQuestion(queue.current[0], allCards));
  const [selected, setSelected] = useState(null);
  const [showHint, setShowHint] = useState(false);
  const [correctCount, setCorrectCount] = useState(0);
  const [updates, setUpdates] = useState([]);
  const [xp, setXp] = useState(0);

  const total = queue.current.length;

  const recordAnswer = (isCorrect, choiceId = question.card.id, skipped = false) => {
    if (selected) return;
    setSelected({ choiceId, isCorrect, skipped });

    const now = Date.now();
    const prevLevel = (initialProgress[question.card.id] || emptyProgress()).level;
    const alreadyUpdated = updates.find((u) => u.id === question.card.id);
    const baseLevel = alreadyUpdated ? alreadyUpdated.level : prevLevel;
    const newLevel = isCorrect ? Math.min(baseLevel + LEVEL_GAIN_CORRECT, MAX_LEVEL) : Math.max(baseLevel - LEVEL_LOSS_WRONG, 0);
    const dueAt = now + reviewIntervalDays(newLevel) * DAY_MS;
    setUpdates((u) => [...u.filter((x) => x.id !== question.card.id), { id: question.card.id, level: newLevel, dueAt, lastReview: now }]);
    if (isCorrect) {
      playCorrectSound();
      speakText(getCardAnswerText(question.card));
      setCorrectCount((n) => n + 1);
      setXp((x) => x + 10);
    }
  };
  const handleChoice = (choice) => {
    recordAnswer(choice.id === question.card.id, choice.id);
  };
  const handleSkip = () => {
    recordAnswer(false, "skip", true);
  };

  const handleNext = () => {
    if (typeof window !== "undefined" && window.speechSynthesis) window.speechSynthesis.cancel();
    const nextIndex = qIndex + 1;
    if (nextIndex >= total) { onFinish({ total, correct: correctCount, xpEarned: xp, updates }); return; }
    setQIndex(nextIndex);
    setQuestion(makeQuestion(queue.current[nextIndex], allCards));
    setSelected(null);
    setShowHint(false);
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
        <>
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
          {(q.card.hint || q.card.hintImage) && (
            <div className="mt-3 text-center">
              <button
                type="button"
                onClick={() => setShowHint((value) => !value)}
                className="rounded border border-[#73bfd7] bg-white px-3 py-1.5 text-xs font-bold text-[#166078] hover:bg-[#e8f7fb]"
              >
                ヒント
              </button>
              {showHint && (
                <div className="mx-auto mt-3 max-w-sm rounded border border-[#b7d6e6] bg-white/80 p-3 text-left text-xs leading-5 text-[#42677a]">
                  <HintContent text={q.card.hint} image={q.card.hintImage} />
                </div>
              )}
            </div>
          )}
        </>
      )}

      {!selected && (
        <button
          type="button"
          onClick={handleSkip}
          className="mt-4 w-full rounded-md border border-[#b7d6e6] bg-white/70 py-2 font-display text-sm font-semibold text-[#42677a] hover:bg-[#e8f7fb]"
        >
          スキップ
        </button>
      )}

      {selected && (
        <div className="mt-5 slide-up">
          <div className={`text-center font-display font-bold text-lg mb-3 ${selected.isCorrect ? "text-[#16805d]" : "text-[#b42335]"}`}>
            {selected.isCorrect ? "✓ 正解！" : selected.skipped ? `スキップ: ${getCardAnswerText(q.card)}` : q.type === "pic2en" ? `✗ ${q.card.en}` : "✗"}
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
function Summary({ result, onReviewAgain, onHome }) {
  const pct = result.total > 0 ? Math.round((result.correct / result.total) * 100) : 0;
  return (
    <div className="pt-10 text-center slide-up">
      <div className="text-6xl mb-3">{pct >= 80 ? "🏅" : pct >= 50 ? "📗" : "📖"}</div>
      <h2 className="font-display text-2xl font-bold text-[#16475f] mb-1">セッション終了</h2>
      <p className="text-[#1687a7] mb-6 font-mono text-sm">{result.correct} / {result.total} 正解（{pct}%）・ +{result.xpEarned} XP</p>
      {result.levelDelta > 0 && <p className="-mt-4 mb-6 font-display text-sm font-bold text-[#16805d]">Lv +{result.levelDelta}</p>}
      {result.levelDelta < 0 && <p className="-mt-4 mb-6 font-display text-sm font-bold text-[#b42335]">Lv {result.levelDelta}</p>}
      <div className="flex flex-col gap-3 max-w-xs mx-auto">
        <button onClick={onReviewAgain} className="rounded-md py-3 flex items-center justify-center gap-2 bg-[#16805d] text-[#ffffff] font-display font-bold hover:brightness-110 transition"><RotateCcw size={18} /> もう一度復習する</button>
        <button onClick={onHome} className="rounded-md py-3 flex items-center justify-center gap-2 bg-[#1687a7] text-[#ffffff] font-display font-semibold hover:brightness-110 transition"><HomeIcon size={16} /> レッスン一覧へ</button>
      </div>
    </div>
  );
}
