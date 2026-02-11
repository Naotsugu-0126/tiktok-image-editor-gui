import express from 'express';
import fs from 'node:fs';
import path from 'node:path';
import {promises as fsp} from 'node:fs';
import {createHash, randomBytes} from 'node:crypto';
import {spawn} from 'node:child_process';
import {pipeline} from 'node:stream/promises';
import {Readable} from 'node:stream';
import {fileURLToPath} from 'node:url';
import {google} from 'googleapis';
import {bundle} from '@remotion/bundler';
import {renderStill, selectComposition} from '@remotion/renderer';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);
const PROJECT_ROOT = path.resolve(__dirname, '..');
const PUBLIC_DIR = path.resolve(__dirname, 'public');
const APP_DATA_DIR = path.resolve(PROJECT_ROOT, 'app-data');
const LOCAL_IMAGE_DIR = path.resolve(APP_DATA_DIR, 'local-images');
const LOCAL_CONFIG_PATH = path.resolve(APP_DATA_DIR, 'local-config.json');
const LOCAL_META_PATH = path.resolve(APP_DATA_DIR, 'local-story-meta.json');
const DEFAULT_TOKEN_PATH = path.resolve(PROJECT_ROOT, '..', 'google_ops_setup', 'oauth_token.json');
const CSV_PATH = path.resolve(PROJECT_ROOT, 'templates', 'captions.csv');
const SOZAI_DIR = path.resolve(PROJECT_ROOT, 'sozai');
const OUTPUT_DIR = path.resolve(PROJECT_ROOT, 'output');
const ENTRY_POINT = path.resolve(PROJECT_ROOT, 'src', 'index.ts');
const PORT = Number.parseInt(process.env.APP_PORT ?? '4173', 10);
const NPM_COMMAND = process.platform === 'win32' ? 'npm.cmd' : 'npm';
const PREVIEW_SCALE = 0.5;
const PREVIEW_JPEG_QUALITY = 42;
const PREVIEW_MODE = 'fast';
const PREVIEW_CACHE_VERSION = 'v1';
const GOOGLE_OAUTH_SCOPES = Object.freeze([
  'https://www.googleapis.com/auth/spreadsheets',
  'https://www.googleapis.com/auth/drive',
]);
const AUTH_STATE_TTL_MS = 10 * 60 * 1000;
const LOCAL_IMAGE_REF_PREFIX = 'local:';
const BATCH_MAX_STORIES = 50;
const SCENES_PER_STORY = 20;

const imageMimeTypes = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
};

const TELOP_TYPES = Object.freeze(['標準', '帯', '吹き出し', '角丸ラベル', '強調', 'ネオン', 'ガラス', 'アウトライン']);
const TEXT_POSITIONS = Object.freeze(['上', '中央', '下']);

const STATUS = Object.freeze({
  draft: '下書き',
  pending: '待機中',
  rendering: '書き出し中',
  done: '完了',
  error: 'エラー',
});

const STATUS_LIST = Object.freeze([STATUS.draft, STATUS.pending, STATUS.rendering, STATUS.done, STATUS.error]);

const STORY_HEADERS = ['ストーリーID', 'タイトル', '素材画像ID'];

const SCENE_HEADERS = ['ストーリーID', 'シーン番号', '本文（入力）'];

const STORY_TEXT_HEADERS = Object.freeze([...STORY_HEADERS]);
const SCENE_TEXT_HEADERS = Object.freeze([...SCENE_HEADERS]);

const GUIDE_ROWS = [
  [
    '使い方',
    '運用方法',
    'スプレッドシートのテキスト管理のみ。詳細設定（テーマ/テロップ/位置/レンダリング）はGUIで調整',
    'シートは A〜C 列のみ',
  ],
  [
    'バッチ上限',
    '最大件数',
    `1バッチ最大${BATCH_MAX_STORIES}ストーリー（1ストーリー${SCENES_PER_STORY}シーン）`,
    `シーン一覧は最大${BATCH_MAX_STORIES * SCENES_PER_STORY}行`,
  ],
  ['素材画像ID', '入力形式', 'DriveファイルID or local:xxxx.png', '1aBcDe... / local:sample_123.png'],
  ['状態', '管理方法', '状態は自動更新。シートで手入力しない', '待機中/完了などはシステム更新'],
  ['テロップ種類', '初期値', '標準 / 帯 / 吹き出し / 角丸ラベル / 強調 / ネオン / ガラス / アウトライン', '標準'],
  ['テキスト位置', '初期値', '上 / 中央 / 下', '中央'],
];

const SHEETS = Object.freeze({
  stories: {
    title: 'ストーリー一覧',
    headers: STORY_TEXT_HEADERS,
    columns: 3,
  },
  scenes: {
    title: 'シーン一覧',
    headers: SCENE_TEXT_HEADERS,
    columns: 3,
  },
});

const STORY_INPUT_MAP = Object.freeze([
  {
    column: 'A',
    header: STORY_HEADERS[0],
    required: true,
    rule: '必須。重複しないID。シーン一覧A列と同じ値を使う',
    example: 'STORY_20260209_0001',
  },
  {
    column: 'B',
    header: STORY_HEADERS[1],
    required: true,
    rule: '必須。管理しやすいタイトル',
    example: '恋愛ストーリー_01',
  },
  {
    column: 'C',
    header: STORY_HEADERS[2],
    required: true,
    rule: '必須。Drive fileId または local:xxxx.png',
    example: 'local:sample_001.png',
  },
]);

const SCENE_INPUT_MAP = Object.freeze([
  {
    column: 'A',
    header: SCENE_HEADERS[0],
    required: true,
    rule: '必須。ストーリー一覧A列のIDを入れる',
    example: 'STORY_20260209_0001',
  },
  {
    column: 'B',
    header: SCENE_HEADERS[1],
    required: true,
    rule: `必須。1〜${SCENES_PER_STORY} の連番`,
    example: '1',
  },
  {
    column: 'C',
    header: SCENE_HEADERS[2],
    required: true,
    rule: '必須。表示する本文',
    example: 'ここに1シーン目の本文を入力',
  },
]);

const DEFAULT_THEMES = Object.freeze([
  ['clean', '標準', '#FFFFFF', '#111111', 'soft', 'none'],
  ['pop', 'ゴシック', '#FFFFFF', '#1B1B1B', 'bold', 'light'],
  ['cinematic', '明朝', '#F8F8F2', '#000000', 'deep', 'dark'],
  ['noir', '明朝', '#F2F2F2', '#0A0A0A', 'deep', 'dark'],
  ['sunset', '太ゴシック', '#FFF7EE', '#4A1E14', 'bold', 'warm'],
  ['aqua', '丸文字', '#F4FFFF', '#003A46', 'soft', 'cool'],
]);

const DEFAULT_RENDER_STYLE = Object.freeze({
  theme: 'clean',
  telopType: TELOP_TYPES[0] || '',
  textPosition: TEXT_POSITIONS[1] || TEXT_POSITIONS[0] || '',
  font: 'standard',
  color: '#FFFFFF',
  strokeColor: '#121212',
  strokeWidth: 0.8,
});
const FONT_STYLE_OPTIONS = Object.freeze(['標準', 'ゴシック', '明朝', '太ゴシック', '細ゴシック', 'UDゴシック', '丸文字', 'standard']);

const normalizeHexColor = (value, fallback = '#FFFFFF') => {
  const text = String(value ?? '').trim();
  if (/^#[0-9a-fA-F]{6}$/.test(text)) {
    return text.toUpperCase();
  }
  return fallback;
};

const normalizeStrokeWidth = (value, fallback = 0.8) => {
  const parsed = Number.parseFloat(String(value ?? ''));
  if (!Number.isFinite(parsed)) {
    return fallback;
  }
  const clamped = Math.max(0, Math.min(4, parsed));
  return Math.round(clamped * 10) / 10;
};

const normalizeRenderStyle = (value) => {
  const raw = value && typeof value === 'object' ? value : {};
  const theme = String(raw.theme ?? '').trim() || DEFAULT_RENDER_STYLE.theme;
  const telopTypeRaw = String(raw.telopType ?? '').trim();
  const textPositionRaw = String(raw.textPosition ?? '').trim();
  const fontRaw = String(raw.font ?? '').trim();
  const font = FONT_STYLE_OPTIONS.includes(fontRaw) ? fontRaw : DEFAULT_RENDER_STYLE.font;
  return {
    theme,
    telopType: TELOP_TYPES.includes(telopTypeRaw) ? telopTypeRaw : DEFAULT_RENDER_STYLE.telopType,
    textPosition: TEXT_POSITIONS.includes(textPositionRaw) ? textPositionRaw : DEFAULT_RENDER_STYLE.textPosition,
    font,
    color: normalizeHexColor(raw.color, DEFAULT_RENDER_STYLE.color),
    strokeColor: normalizeHexColor(raw.strokeColor, DEFAULT_RENDER_STYLE.strokeColor),
    strokeWidth: normalizeStrokeWidth(raw.strokeWidth, DEFAULT_RENDER_STYLE.strokeWidth),
  };
};

const jobs = new Map();
let jobQueueTail = Promise.resolve();
const previewRenderLocks = new Map();
const previewImageCache = new Map();
const oauthAuthStates = new Map();
const PREVIEW_IMAGE_CACHE_LIMIT = 24;
const PREVIEW_IMAGE_CACHE_TTL_MS = 5 * 60 * 1000;
let cachedBundleLocationPromise = null;

const asyncHandler = (fn) => (req, res, next) => Promise.resolve(fn(req, res, next)).catch(next);

const setPreviewImageCache = (key, value) => {
  if (!key) return;
  if (previewImageCache.has(key)) {
    previewImageCache.delete(key);
  }
  previewImageCache.set(key, value);
  while (previewImageCache.size > PREVIEW_IMAGE_CACHE_LIMIT) {
    const oldestKey = previewImageCache.keys().next().value;
    if (!oldestKey) break;
    previewImageCache.delete(oldestKey);
  }
};

const ensureAppDataDir = () => {
  fs.mkdirSync(APP_DATA_DIR, {recursive: true});
};

const ensureLocalImageDir = () => {
  fs.mkdirSync(LOCAL_IMAGE_DIR, {recursive: true});
};

const sanitizeLocalFileName = (value) =>
  String(value ?? '')
    .trim()
    .replace(/[<>:"/\\|?*]/g, '-')
    .replace(/\s+/g, '_');

const extFromUploadedMime = (mime) => {
  if (mime === 'image/png') return '.png';
  if (mime === 'image/jpeg') return '.jpg';
  if (mime === 'image/webp') return '.webp';
  throw new Error(`譛ｪ蟇ｾ蠢懊・逕ｻ蜒丞ｽ｢蠑上〒縺・ ${mime}`);
};

const parseImageDataUrl = (dataUrl) => {
  const raw = String(dataUrl ?? '').trim();
  const match = raw.match(/^data:(image\/(?:png|jpeg|webp));base64,([A-Za-z0-9+/=]+)$/);
  if (!match) {
    throw new Error('Required data is missing.');
  }
  return {
    mimeType: match[1],
    buffer: Buffer.from(match[2], 'base64'),
  };
};

const isLocalImageRef = (value) => String(value ?? '').startsWith(LOCAL_IMAGE_REF_PREFIX);

const resolveLocalImagePathFromRef = (imageRef) => {
  const raw = String(imageRef ?? '').trim();
  if (!isLocalImageRef(raw)) {
    throw new Error(`繝ｭ繝ｼ繧ｫ繝ｫ逕ｻ蜒丞盾辣ｧ縺ｧ縺ｯ縺ゅｊ縺ｾ縺帙ｓ: ${raw}`);
  }
  const fileName = raw.slice(LOCAL_IMAGE_REF_PREFIX.length).trim();
  const baseName = path.basename(fileName);
  if (!baseName || baseName !== fileName) {
    throw new Error(`繝ｭ繝ｼ繧ｫ繝ｫ逕ｻ蜒丞盾辣ｧ縺御ｸ肴ｭ｣縺ｧ縺・ ${raw}`);
  }
  return path.resolve(LOCAL_IMAGE_DIR, baseName);
};

const readJsonIfExists = (targetPath, fallback) => {
  if (!fs.existsSync(targetPath)) {
    return fallback;
  }
  try {
    return JSON.parse(fs.readFileSync(targetPath, 'utf8'));
  } catch {
    return fallback;
  }
};

const readLocalConfig = () => {
  ensureAppDataDir();
  const stored = readJsonIfExists(LOCAL_CONFIG_PATH, {});
  return {
    spreadsheetId: String(stored.spreadsheetId ?? '').trim(),
    tokenPath: String(stored.tokenPath ?? '').trim(),
    outputDriveFolderId: String(stored.outputDriveFolderId ?? '').trim(),
    renderStyle: normalizeRenderStyle(stored.renderStyle),
  };
};

const readLocalMeta = () => {
  ensureAppDataDir();
  const stored = readJsonIfExists(LOCAL_META_PATH, {});
  const stories = stored && typeof stored === 'object' && stored.stories && typeof stored.stories === 'object'
    ? stored.stories
    : {};
  return {stories};
};

const writeLocalMeta = (next) => {
  ensureAppDataDir();
  fs.writeFileSync(LOCAL_META_PATH, `${JSON.stringify(next, null, 2)}\n`, 'utf8');
};

const ensureStoryMeta = (meta, storyId) => {
  if (!meta.stories[storyId] || typeof meta.stories[storyId] !== 'object') {
    meta.stories[storyId] = {};
  }
  const storyMeta = meta.stories[storyId];
  if (!storyMeta.scenes || typeof storyMeta.scenes !== 'object') {
    storyMeta.scenes = {};
  }
  return storyMeta;
};

const getRuntimeConfig = () => {
  const local = readLocalConfig();
  const envSpreadsheetId = String(process.env.GOOGLE_SHEET_ID ?? '').trim();
  const envTokenPath = String(process.env.GOOGLE_TOKEN_PATH ?? '').trim();
  const envOutputDriveFolderId = String(process.env.GOOGLE_DRIVE_OUTPUT_FOLDER_ID ?? '').trim();
  return {
    spreadsheetId: envSpreadsheetId || local.spreadsheetId,
    tokenPath: envTokenPath || local.tokenPath || DEFAULT_TOKEN_PATH,
    outputDriveFolderId: envOutputDriveFolderId || local.outputDriveFolderId,
    renderStyle: normalizeRenderStyle(local.renderStyle),
  };
};

const writeLocalConfig = (patch) => {
  ensureAppDataDir();
  const current = readLocalConfig();
  const next = {
    ...current,
    ...patch,
  };
  next.renderStyle = normalizeRenderStyle(next.renderStyle);
  fs.writeFileSync(LOCAL_CONFIG_PATH, `${JSON.stringify(next, null, 2)}\n`, 'utf8');
  return next;
};

const nowIso = () => new Date().toISOString();

const normalizeHeader = (value) =>
  String(value ?? '')
    .trim()
    .replace(/[()・茨ｼ噂s]/g, '')
    .toLowerCase();

const getBundleLocation = async () => {
  if (!cachedBundleLocationPromise) {
    cachedBundleLocationPromise = bundle({
      entryPoint: ENTRY_POINT,
      webpackOverride: (config) => config,
    });
  }
  return cachedBundleLocationPromise;
};

const withPreviewLock = async (key, task) => {
  const existing = previewRenderLocks.get(key);
  if (existing) {
    return existing;
  }

  const current = task().finally(() => {
    previewRenderLocks.delete(key);
  });
  previewRenderLocks.set(key, current);
  return current;
};

const escapeSheetTitle = (title) => title.replace(/'/g, "''");

const columnNameFromNumber = (n) => {
  let num = n;
  let out = '';
  while (num > 0) {
    const mod = (num - 1) % 26;
    out = String.fromCharCode(65 + mod) + out;
    num = Math.floor((num - mod) / 26);
  }
  return out;
};

const buildSheetRange = (sheetTitle, from, to) =>
  `'${escapeSheetTitle(sheetTitle)}'!${from}:${to}`;

const parseAppendRowNumber = (updatedRange) => {
  const match = String(updatedRange ?? '').match(/![A-Z]+(\d+)(?::[A-Z]+(\d+))?/);
  if (!match) return null;
  return Number.parseInt(match[1], 10);
};

const buildSpreadsheetUrl = (spreadsheetId) => {
  const id = String(spreadsheetId ?? '').trim();
  return id ? `https://docs.google.com/spreadsheets/d/${id}/edit` : '';
};

const fetchWithTimeout = async (url, options = {}, timeoutMs = 15000) => {
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), timeoutMs);
  try {
    return await fetch(url, {
      ...options,
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timer);
  }
};

const parseGvizPayload = (rawText) => {
  const text = String(rawText ?? '');
  const start = text.indexOf('{');
  const end = text.lastIndexOf('}');
  if (start < 0 || end <= start) {
    throw new Error('Could not parse public spreadsheet response.');
  }
  const payload = JSON.parse(text.slice(start, end + 1));
  const status = String(payload?.status ?? 'ok').toLowerCase();
  if (status !== 'ok') {
    const detail =
      String(payload?.errors?.[0]?.detailed_message ?? '').trim() ||
      String(payload?.errors?.[0]?.message ?? '').trim() ||
      'Failed to load public spreadsheet.';
    throw new Error(detail);
  }
  return payload;
};

const toGvizCellString = (cell) => {
  if (!cell || typeof cell !== 'object') return '';
  const value = cell.v;
  if (value === null || value === undefined) {
    return typeof cell.f === 'string' ? cell.f : '';
  }
  if (typeof value === 'string') return value;
  if (typeof value === 'number' || typeof value === 'boolean') return String(value);
  if (typeof cell.f === 'string') return cell.f;
  return String(value);
};

const fetchPublicSheetRows = async (spreadsheetId, sheetTitle) => {
  const id = String(spreadsheetId ?? '').trim();
  if (!id) {
    throw new Error('Spreadsheet ID is not configured.');
  }
  const title = String(sheetTitle ?? '').trim();
  const url =
    `https://docs.google.com/spreadsheets/d/${encodeURIComponent(id)}/gviz/tq` +
    `?sheet=${encodeURIComponent(title)}&tqx=out:json`;
  const res = await fetchWithTimeout(url, {method: 'GET'}, 15000);
  if (!res.ok) {
    throw new Error(`蜈ｬ髢九せ繝励Ξ繝・ラ繧ｷ繝ｼ繝医・蜿門ｾ励↓螟ｱ謨励＠縺ｾ縺励◆ (HTTP ${res.status})`);
  }
  const payload = parseGvizPayload(await res.text());
  const rows = Array.isArray(payload?.table?.rows) ? payload.table.rows : [];
  return rows.map((row) => {
    const cells = Array.isArray(row?.c) ? row.c : [];
    return cells.map((cell) => toGvizCellString(cell));
  });
};

const parseSpreadsheetIdInput = (value) => {
  const raw = String(value ?? '').trim();
  if (!raw) return '';
  const match = raw.match(/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
  if (match?.[1]) {
    return match[1];
  }
  return raw;
};

const getOAuthRedirectUri = () => `http://127.0.0.1:${PORT}/api/auth/google/callback`;

const readTokenJsonIfExists = (tokenPath) => {
  try {
    if (!tokenPath || !fs.existsSync(tokenPath)) {
      return null;
    }
    const raw = fs.readFileSync(tokenPath, 'utf8');
    return JSON.parse(raw);
  } catch {
    return null;
  }
};

const resolveOAuthClientInfo = (config) => {
  const tokenPath = String(config?.tokenPath || '').trim() || DEFAULT_TOKEN_PATH;
  const tokenJson = readTokenJsonIfExists(tokenPath);
  const fallbackTokenJson =
    tokenPath === DEFAULT_TOKEN_PATH ? tokenJson : tokenJson || readTokenJsonIfExists(DEFAULT_TOKEN_PATH);
  const envClientId = String(process.env.GOOGLE_CLIENT_ID ?? '').trim();
  const envClientSecret = String(process.env.GOOGLE_CLIENT_SECRET ?? '').trim();
  const clientId = envClientId || String(tokenJson?.client_id ?? fallbackTokenJson?.client_id ?? '').trim();
  const clientSecret =
    envClientSecret || String(tokenJson?.client_secret ?? fallbackTokenJson?.client_secret ?? '').trim();
  const hasRefreshToken = Boolean(String(tokenJson?.refresh_token ?? '').trim());
  return {
    tokenPath,
    tokenJson,
    clientId,
    clientSecret,
    hasRefreshToken,
    ready: Boolean(clientId && clientSecret),
  };
};

const buildOAuthClient = ({clientId, clientSecret, redirectUri}) =>
  new google.auth.OAuth2(clientId, clientSecret, redirectUri || getOAuthRedirectUri());

const pruneOAuthAuthStates = () => {
  const now = Date.now();
  for (const [state, entry] of oauthAuthStates.entries()) {
    if (!entry || now - Number(entry.createdAt || 0) > AUTH_STATE_TTL_MS) {
      oauthAuthStates.delete(state);
    }
  }
};

const saveOAuthToken = async ({tokenPath, clientId, clientSecret, redirectUri, tokens}) => {
  const existing = readTokenJsonIfExists(tokenPath) || {};
  const merged = {
    ...existing,
    ...tokens,
    client_id: clientId,
    client_secret: clientSecret,
    redirect_uri: redirectUri || getOAuthRedirectUri(),
    refresh_token: tokens.refresh_token ?? existing.refresh_token,
  };
  const targetDir = path.dirname(tokenPath);
  await fsp.mkdir(targetDir, {recursive: true});
  await fsp.writeFile(tokenPath, `${JSON.stringify(merged, null, 2)}\n`, 'utf8');
  return merged;
};

const escapeHtmlText = (value) =>
  String(value ?? '')
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;');

const renderOAuthPopupPage = ({ok, message, payload}) => {
  const safeMessage = escapeHtmlText(message);
  const payloadJson = JSON.stringify(payload ?? {type: 'google-auth-complete', ok: false});
  return `<!doctype html>
<html lang="ja">
  <head>
    <meta charset="utf-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Google隱崎ｨｼ</title>
    <style>
      body { font-family: 'Yu Gothic UI', Meiryo, sans-serif; margin: 24px; color: #1f2a3a; }
      .ok { color: #1f8f4d; }
      .ng { color: #c61f3d; }
      .box { border: 1px solid #d9e1ec; border-radius: 12px; padding: 14px; background: #fff; }
      .muted { color: #617287; font-size: 13px; margin-top: 8px; }
    </style>
  </head>
  <body>
    <div class="box">
      <h3 class="${ok ? 'ok' : 'ng'}">${ok ? 'Google login complete' : 'Google login failed'}</h3>
      <div>${safeMessage}</div>
      <div class="muted">This window closes automatically. If not, close it manually.</div>
    </div>
    <script>
      (function () {
        try {
          var payload = ${payloadJson};
          if (window.opener && !window.opener.closed) {
            window.opener.postMessage(payload, '*');
          }
        } catch (_e) {}
        setTimeout(function () {
          window.close();
        }, 1200);
      })();
    </script>
  </body>
</html>`;
};

const buildInputMapRows = () => [
  ...STORY_INPUT_MAP.map((entry) => ({
    sheet: SHEETS.stories.title,
    column: entry.column,
    header: entry.header,
    required: entry.required,
    rule: entry.rule,
    example: entry.example,
  })),
  ...SCENE_INPUT_MAP.map((entry) => ({
    sheet: SHEETS.scenes.title,
    column: entry.column,
    header: entry.header,
    required: entry.required,
    rule: entry.rule,
    example: entry.example,
  })),
];

const toGuideSheetRow = (entry) => [
  entry.sheet,
  `${entry.column}列 ${entry.header}${entry.required ? ' (必須)' : ''}`,
  entry.rule,
  entry.example,
];

const ensureFileExists = async (targetPath, name) => {
  try {
    await fsp.access(targetPath, fs.constants.R_OK);
  } catch {
    throw new Error(`${name} 縺瑚ｦ九▽縺九ｊ縺ｾ縺帙ｓ: ${targetPath}`);
  }
};

const findFirstImageFile = async (dir) => {
  const entries = await fsp.readdir(dir);
  const imageEntries = entries
    .filter((name) => imageMimeTypes[path.extname(name).toLowerCase()])
    .sort((a, b) => a.localeCompare(b, 'ja'));
  const candidates = imageEntries.filter(
    (name) => !name.startsWith('000-runtime-') && !name.startsWith('000-preview-')
  );
  const targetList = candidates.length > 0 ? candidates : imageEntries;
  if (targetList.length === 0) {
    throw new Error('Required data is missing.');
  }
  return path.resolve(dir, targetList[0]);
};

const toDataUrl = async (imagePath) => {
  const ext = path.extname(imagePath).toLowerCase();
  const mimeType = imageMimeTypes[ext];
  if (!mimeType) {
    throw new Error(`譛ｪ蟇ｾ蠢懊・逕ｻ蜒丞ｽ｢蠑上〒縺・ ${ext}`);
  }
  const buffer = await fsp.readFile(imagePath);
  return `data:${mimeType};base64,${buffer.toString('base64')}`;
};

const buildGoogleClients = async (config) => {
  await ensureFileExists(config.tokenPath, 'OAuth token file');
  const tokenRaw = await fsp.readFile(config.tokenPath, 'utf8');
  const token = JSON.parse(tokenRaw);

  if (!token.client_id || !token.client_secret) {
    throw new Error('oauth_token.json is missing client_id/client_secret.');
  }

  const oauth2Client = new google.auth.OAuth2(
    token.client_id,
    token.client_secret,
    token.redirect_uri ?? 'http://localhost'
  );
  oauth2Client.setCredentials(token);

  oauth2Client.on('tokens', (tokens) => {
    const merged = {
      ...token,
      ...tokens,
      refresh_token: tokens.refresh_token ?? token.refresh_token,
    };
    fs.writeFileSync(config.tokenPath, `${JSON.stringify(merged, null, 2)}\n`, 'utf8');
  });

  return {
    sheets: google.sheets({version: 'v4', auth: oauth2Client}),
    drive: google.drive({version: 'v3', auth: oauth2Client}),
  };
};

const tryBuildGoogleClients = async (config) => {
  try {
    const clients = await buildGoogleClients(config);
    return {
      mode: 'auth',
      sheets: clients.sheets,
      drive: clients.drive,
      error: null,
    };
  } catch (error) {
    if (!isAuthRecoverableError(error)) {
      throw error;
    }
    return {
      mode: 'public',
      sheets: null,
      drive: null,
      error,
    };
  }
};

const isAuthRecoverableError = (error) => {
  const message = String(error?.message ?? error ?? '').toLowerCase();
  if (!message) return false;
  return (
    message.includes('oauth') ||
    message.includes('token') ||
    message.includes('invalid_grant') ||
    message.includes('invalid_client') ||
    message.includes('unauthorized_client') ||
    message.includes('unauthenticated')
  );
};

const ensureSpreadsheetStructure = async (sheets, spreadsheetId) => {
  const metadata = await sheets.spreadsheets.get({
    spreadsheetId,
    fields: 'properties(title),sheets(properties(title))',
  });

  const existingTitles = new Set(
    (metadata.data.sheets ?? [])
      .map((entry) => entry.properties?.title)
      .filter((title) => typeof title === 'string')
  );

  const missingRequests = Object.values(SHEETS)
    .filter((sheet) => !existingTitles.has(sheet.title))
    .map((sheet) => ({
      addSheet: {
        properties: {
          title: sheet.title,
        },
      },
    }));

  if (missingRequests.length > 0) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: {
        requests: missingRequests,
      },
    });
  }

  for (const sheet of Object.values(SHEETS)) {
    const headerRes = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: buildSheetRange(sheet.title, '1', '1'),
    });
    const row = headerRes.data.values?.[0] ?? [];
    const existing = row.map((value) => normalizeHeader(value));
    const required = sheet.headers.map((value) => normalizeHeader(value));
    let needsUpdate = row.length === 0 || existing.length < required.length;
    if (!needsUpdate) {
      for (let i = 0; i < required.length; i += 1) {
        if (existing[i] !== required[i]) {
          needsUpdate = true;
          break;
        }
      }
    }

    if (needsUpdate) {
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: buildSheetRange(sheet.title, 'A1', `${columnNameFromNumber(sheet.columns)}1`),
        valueInputOption: 'RAW',
        requestBody: {
          values: [sheet.headers],
        },
      });
    }
  }
};

const getSheetRowsWithRowNumber = async (sheets, spreadsheetId, sheet) => {
  if (sheets) {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: buildSheetRange(sheet.title, 'A2', columnNameFromNumber(sheet.columns)),
    });
    const rows = res.data.values ?? [];
    return rows.map((values, index) => ({values, rowNumber: index + 2}));
  }

  const rawRows = await fetchPublicSheetRows(spreadsheetId, sheet.title);
  const normalizedRows = rawRows.map((row) => {
    const out = [];
    for (let i = 0; i < sheet.columns; i += 1) {
      out.push(String(row[i] ?? ''));
    }
    return out;
  });

  const requiredHeaders = sheet.headers.map((value) => normalizeHeader(value));
  const firstRowHeaders = (normalizedRows[0] ?? []).map((value) => normalizeHeader(value));
  const hasHeader =
    normalizedRows.length > 0 &&
    requiredHeaders.every((header, index) => firstRowHeaders[index] === header);
  const startIndex = hasHeader ? 1 : 0;

  return normalizedRows
    .slice(startIndex)
    .map((values, index) => ({values, rowNumber: index + 1 + startIndex}));
};

const normalizeStory = (row, rowNumber, metaStory = null) => ({
  storyId: String(row[0] ?? '').trim(),
  title: String(row[1] ?? '').trim(),
  imageFileId: String(row[2] ?? '').trim(),
  theme: String(metaStory?.theme ?? '').trim() || 'clean',
  status: String(metaStory?.status ?? '').trim() || STATUS.draft,
  createdAt: String(metaStory?.createdAt ?? '').trim(),
  _rowNumber: rowNumber,
});

const normalizeScene = (row, rowNumber, metaScene = null, storyTheme = 'clean') => {
  const sceneNo = Number.parseInt(String(row[1] ?? ''), 10) || 0;
  return {
    storyId: String(row[0] ?? '').trim(),
    sceneNo,
    textRaw: String(row[2] ?? ''),
    textFinal: String(metaScene?.textFinal ?? ''),
    durationSec: Number.parseFloat(String(metaScene?.durationSec ?? '3')) || 3,
    order: Number.parseInt(String(metaScene?.order ?? sceneNo), 10) || sceneNo,
    status: String(metaScene?.status ?? '').trim() || STATUS.draft,
    theme: String(metaScene?.theme ?? '').trim() || String(storyTheme || 'clean'),
    telopType: String(metaScene?.telopType ?? '').trim() || TELOP_TYPES[0] || '標準',
    textPosition: String(metaScene?.textPosition ?? '').trim() || TEXT_POSITIONS[1] || '中央',
    font: String(metaScene?.font ?? '').trim() || '標準',
    color: String(metaScene?.color ?? '').trim() || '#FFFFFF',
    strokeColor: normalizeHexColor(metaScene?.strokeColor, '#121212'),
    strokeWidth: normalizeStrokeWidth(metaScene?.strokeWidth, 0.8),
    outputFileName:
      String(metaScene?.outputFileName ?? '').trim() ||
      `${String(row[0] ?? '').trim()}_${String(sceneNo).padStart(2, '0')}.png`,
    _rowNumber: rowNumber,
  };
};

const normalizeTheme = (row) => ({
  themeName: String(row[0] ?? '').trim(),
  fontFamily: String(row[1] ?? '').trim(),
  textColor: String(row[2] ?? '').trim(),
  strokeColor: String(row[3] ?? '').trim(),
  shadow: String(row[4] ?? '').trim(),
  backgroundBand: String(row[5] ?? '').trim(),
});

const getStories = async (sheets, spreadsheetId) => {
  const rows = await getSheetRowsWithRowNumber(sheets, spreadsheetId, SHEETS.stories);
  const meta = readLocalMeta();
  return rows
    .map((entry) => {
      const storyId = String(entry.values?.[0] ?? '').trim();
      const metaStory = storyId ? meta.stories[storyId] : null;
      return normalizeStory(entry.values, entry.rowNumber, metaStory);
    })
    .filter((story) => story.storyId.length > 0);
};

const getThemes = async () =>
  DEFAULT_THEMES.map((row) => normalizeTheme(row)).filter((theme) => theme.themeName.length > 0);

const getAllScenes = async (sheets, spreadsheetId) => {
  const rows = await getSheetRowsWithRowNumber(sheets, spreadsheetId, SHEETS.scenes);
  const stories = await getStories(sheets, spreadsheetId);
  const storyThemeMap = new Map(stories.map((story) => [story.storyId, story.theme || 'clean']));
  const meta = readLocalMeta();
  return rows
    .map((entry) => {
      const storyId = String(entry.values?.[0] ?? '').trim();
      const sceneNo = Number.parseInt(String(entry.values?.[1] ?? ''), 10) || 0;
      const metaScene = meta.stories?.[storyId]?.scenes?.[String(sceneNo)] ?? null;
      return normalizeScene(entry.values, entry.rowNumber, metaScene, storyThemeMap.get(storyId) || 'clean');
    })
    .filter((scene) => scene.storyId.length > 0);
};

const getStoryDetail = async (sheets, spreadsheetId, storyId) => {
  const stories = await getStories(sheets, spreadsheetId);
  const story = stories.find((entry) => entry.storyId === storyId);
  if (!story) {
    throw new Error(`Story ID not found: ${storyId}`);
  }

  const allScenes = await getAllScenes(sheets, spreadsheetId);
  const storyScenes = allScenes.filter((scene) => scene.storyId === storyId);
  const bySceneNo = new Map(storyScenes.map((scene) => [scene.sceneNo, scene]));

  const scenes = [];
  for (let i = 1; i <= SCENES_PER_STORY; i += 1) {
    const existing = bySceneNo.get(i);
    if (existing) {
      scenes.push(existing);
    } else {
      scenes.push({
        storyId,
        sceneNo: i,
        textRaw: '',
        textFinal: '',
        durationSec: 3,
        order: i,
        status: STATUS.draft,
        theme: story.theme || 'clean',
        telopType: '',
        textPosition: '',
        font: '',
        color: '#FFFFFF',
        strokeColor: '#121212',
        strokeWidth: 0.8,
        outputFileName: `${storyId}_${String(i).padStart(2, '0')}.png`,
        _rowNumber: null,
      });
    }
  }

  return {story, scenes};
};

const generateStoryId = () => {
  const date = new Date();
  const pad2 = (value) => String(value).padStart(2, '0');
  const pad3 = (value) => String(value).padStart(3, '0');
  const rand = Math.floor(Math.random() * 1000);
  return `STORY_${date.getFullYear()}${pad2(date.getMonth() + 1)}${pad2(date.getDate())}_${pad2(
    date.getHours()
  )}${pad2(date.getMinutes())}${pad2(date.getSeconds())}${pad3(date.getMilliseconds())}${pad3(rand)}`;
};

const createStory = async (sheets, spreadsheetId, payload) => {
  const stories = await getStories(sheets, spreadsheetId);
  const storyId = String(payload.storyId ?? '').trim() || generateStoryId();
  if (stories.some((entry) => entry.storyId === storyId)) {
    throw new Error(`Story ID already exists: ${storyId}`);
  }

  const title = String(payload.title ?? '').trim();
  if (!title) {
    throw new Error('Required data is missing.');
  }

  const imageFileId = String(payload.imageFileId ?? '').trim();
  const theme = String(payload.theme ?? '').trim() || 'clean';
  const createdAt = nowIso();
  const status = STATUS.draft;

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: buildSheetRange(SHEETS.stories.title, 'A', 'C'),
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: [[storyId, title, imageFileId]],
    },
  });

  const sceneRows = [];
  for (let i = 1; i <= SCENES_PER_STORY; i += 1) {
    sceneRows.push([
      storyId,
      i,
      '',
    ]);
  }

  await sheets.spreadsheets.values.append({
    spreadsheetId,
    range: buildSheetRange(SHEETS.scenes.title, 'A', 'C'),
    valueInputOption: 'USER_ENTERED',
    insertDataOption: 'INSERT_ROWS',
    requestBody: {
      values: sceneRows,
    },
  });

  const meta = readLocalMeta();
  const storyMeta = ensureStoryMeta(meta, storyId);
  storyMeta.theme = theme;
  storyMeta.status = status;
  storyMeta.createdAt = createdAt;
  writeLocalMeta(meta);

  return {
    storyId,
    title,
    imageFileId,
    theme,
    status,
    createdAt,
  };
};

const updateStoryStatus = async (sheets, spreadsheetId, storyId, status) => {
  const _ignore = {sheets, spreadsheetId};
  void _ignore;
  const nextStatus = String(status ?? '').trim() || STATUS.draft;
  const meta = readLocalMeta();
  const storyMeta = ensureStoryMeta(meta, storyId);
  storyMeta.status = nextStatus;
  writeLocalMeta(meta);
};

const updateStoryImageRef = async (sheets, spreadsheetId, storyId, imageRef) => {
  const stories = await getStories(sheets, spreadsheetId);
  const story = stories.find((entry) => entry.storyId === storyId);
  if (!story) {
    throw new Error(`Story ID not found: ${storyId}`);
  }
  await sheets.spreadsheets.values.update({
    spreadsheetId,
    range: buildSheetRange(SHEETS.stories.title, `C${story._rowNumber}`, `C${story._rowNumber}`),
    valueInputOption: 'RAW',
    requestBody: {
      values: [[String(imageRef ?? '').trim()]],
    },
  });
};

const normalizeScenePayload = (storyId, scene) => {
  const sceneNo = Number.parseInt(String(scene.sceneNo ?? ''), 10);
  if (!Number.isFinite(sceneNo) || sceneNo <= 0) {
    throw new Error('Required data is missing.');
  }
  const status = String(scene.status ?? '').trim() || STATUS.draft;
  if (!STATUS_LIST.includes(status)) {
    throw new Error(`迥ｶ諷九′荳肴ｭ｣縺ｧ縺・ ${status}`);
  }
  const theme = String(scene.theme ?? '').trim() || 'clean';
  const telopType = String(scene.telopType ?? '').trim() || TELOP_TYPES[0] || '標準';
  if (!TELOP_TYPES.includes(telopType)) {
    throw new Error(`繝・Ο繝・・遞ｮ鬘槭′荳肴ｭ｣縺ｧ縺・ ${telopType}`);
  }
  const textPosition = String(scene.textPosition ?? '').trim() || TEXT_POSITIONS[1] || '中央';
  if (!TEXT_POSITIONS.includes(textPosition)) {
    throw new Error(`繝・く繧ｹ繝井ｽ咲ｽｮ縺御ｸ肴ｭ｣縺ｧ縺・ ${textPosition}`);
  }
  const font = String(scene.font ?? '').trim() || '標準';
  const color = normalizeHexColor(scene.color, '#FFFFFF');
  const strokeColor = normalizeHexColor(scene.strokeColor, '#121212');
  const strokeWidth = normalizeStrokeWidth(scene.strokeWidth, 0.8);
  const textRaw = String(scene.textRaw ?? '');
  const textFinal = String(scene.textFinal ?? textRaw);
  const durationSec = Number.parseFloat(String(scene.durationSec ?? '3')) || 3;
  const order = Number.parseInt(String(scene.order ?? sceneNo), 10) || sceneNo;
  const outputFileName =
    String(scene.outputFileName ?? '').trim() || `${storyId}_${String(sceneNo).padStart(2, '0')}.png`;

  return {
    storyId,
    sceneNo,
    textRaw,
    textFinal,
    durationSec,
    order,
    status,
    theme,
    telopType,
    textPosition,
    font,
    color,
    strokeColor,
    strokeWidth,
    outputFileName,
  };
};

const buildScenePayloadsFromTemplate = ({
  storyId,
  theme,
  sceneTexts,
  telopType = TELOP_TYPES[0] || '標準',
  textPosition = TEXT_POSITIONS[1] || '中央',
  font = '標準',
  color = '#FFFFFF',
  strokeColor = '#121212',
  strokeWidth = 0.8,
}) => {
  const payloads = [];
  for (let i = 0; i < sceneTexts.length; i += 1) {
    const sceneNo = i + 1;
    const text = String(sceneTexts[i] ?? '');
    payloads.push({
      sceneNo,
      textRaw: text,
      textFinal: text,
      status: text.trim().length > 0 ? STATUS.pending : STATUS.draft,
      theme,
      telopType,
      textPosition,
      outputFileName: `${storyId}_${String(sceneNo).padStart(2, '0')}.png`,
      durationSec: 3,
      order: sceneNo,
      font,
      color,
      strokeColor,
      strokeWidth,
    });
  }
  return payloads;
};

const saveScenes = async (sheets, spreadsheetId, storyId, scenePayloads) => {
  const existingScenes = (await getAllScenes(sheets, spreadsheetId)).filter((scene) => scene.storyId === storyId);
  const bySceneNo = new Map(existingScenes.map((scene) => [scene.sceneNo, scene]));
  const meta = readLocalMeta();
  const storyMeta = ensureStoryMeta(meta, storyId);

  const updates = [];
  const appends = [];
  for (const rawScene of scenePayloads) {
    const scene = normalizeScenePayload(storyId, rawScene);
    const values = [scene.storyId, scene.sceneNo, scene.textRaw];
    const existing = bySceneNo.get(scene.sceneNo);
    if (existing && existing._rowNumber) {
      updates.push({
        range: buildSheetRange(SHEETS.scenes.title, `A${existing._rowNumber}`, `C${existing._rowNumber}`),
        values: [values],
      });
    } else {
      appends.push(values);
    }
    storyMeta.scenes[String(scene.sceneNo)] = {
      textFinal: scene.textFinal,
      durationSec: scene.durationSec,
      order: scene.order,
      status: scene.status,
      theme: scene.theme,
      telopType: scene.telopType,
      textPosition: scene.textPosition,
      font: scene.font,
      color: scene.color,
      strokeColor: scene.strokeColor,
      strokeWidth: scene.strokeWidth,
      outputFileName: scene.outputFileName,
    };
  }

  if (updates.length > 0) {
    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId,
      requestBody: {
        valueInputOption: 'USER_ENTERED',
        data: updates,
      },
    });
  }

  if (appends.length > 0) {
    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: buildSheetRange(SHEETS.scenes.title, 'A', 'C'),
      valueInputOption: 'USER_ENTERED',
      insertDataOption: 'INSERT_ROWS',
      requestBody: {
        values: appends,
      },
    });
  }

  storyMeta.status = STATUS.draft;
  writeLocalMeta(meta);

  return {
    updated: updates.length,
    appended: appends.length,
  };
};

const buildCsv = (rows) => {
  const escape = (value) => {
    const text = String(value ?? '');
    if (text.includes('"') || text.includes(',') || text.includes('\n')) {
      return `"${text.replaceAll('"', '""')}"`;
    }
    return text;
  };
  return rows.map((row) => row.map((field) => escape(field)).join(',')).join('\n');
};

const buildTsv = (rows) => rows.map((row) => row.map((value) => String(value ?? '')).join('\t')).join('\n');

const buildSheetGuidePayload = (config = getRuntimeConfig()) => {
  const storySample = [
    'STORY_001',
    'TikTokサンプル_ストーリー',
    'local:sample_001.png',
  ];
  const sceneSample = [
    'STORY_001',
    '1',
    'ここに1シーン目の本文を入力',
  ];

  const storyTemplateTsv = `${buildTsv([STORY_TEXT_HEADERS, storySample])}\n`;
  const sceneTemplateTsv = `${buildTsv([SCENE_TEXT_HEADERS, sceneSample])}\n`;
  const inputMapRows = buildInputMapRows();

  return {
    storyHeaders: STORY_HEADERS,
    sceneHeaders: SCENE_HEADERS,
    storyTextHeaders: STORY_TEXT_HEADERS,
    sceneTextHeaders: SCENE_TEXT_HEADERS,
    storyTemplateTsv,
    sceneTemplateTsv,
    guideRows: GUIDE_ROWS,
    inputMapRows,
    batchConfig: {
      maxStories: BATCH_MAX_STORIES,
      scenesPerStory: SCENES_PER_STORY,
      maxSceneRows: BATCH_MAX_STORIES * SCENES_PER_STORY,
    },
    sheetNames: {
      stories: SHEETS.stories.title,
      scenes: SHEETS.scenes.title,
    },
    sheetLocation: {
      spreadsheetId: config.spreadsheetId || '',
      spreadsheetUrl: buildSpreadsheetUrl(config.spreadsheetId),
      localConfigPath: LOCAL_CONFIG_PATH,
      localImageDir: LOCAL_IMAGE_DIR,
    },
  };
};

const cleanSozaiFilesByPrefix = async (prefix) => {
  await fsp.mkdir(SOZAI_DIR, {recursive: true});
  const entries = await fsp.readdir(SOZAI_DIR);
  const targets = entries.filter((name) => name.startsWith(prefix));
  await Promise.all(targets.map((name) => fsp.unlink(path.resolve(SOZAI_DIR, name)).catch(() => {})));
};

const extFromMimeType = (mimeType) => {
  switch (mimeType) {
    case 'image/png':
      return '.png';
    case 'image/jpeg':
      return '.jpg';
    case 'image/webp':
      return '.webp';
    default:
      return '.png';
  }
};

const isSupportedImageExt = (ext) => Boolean(imageMimeTypes[String(ext ?? '').toLowerCase()]);

const resolveImageExt = (fileName, mimeType) => {
  const fromName = path.extname(String(fileName ?? '')).toLowerCase();
  if (isSupportedImageExt(fromName)) {
    return fromName;
  }
  return extFromMimeType(String(mimeType ?? ''));
};

const downloadDriveImage = async (drive, fileId, storyId, prefix = '000-runtime-') => {
  await cleanSozaiFilesByPrefix(prefix);
  const metadata = await drive.files.get({
    fileId,
    fields: 'name,mimeType,modifiedTime',
  });

  const fileName = String(metadata.data.name ?? '');
  const ext = resolveImageExt(fileName, metadata.data.mimeType);
  const destName = `${prefix}${storyId}${ext}`;
  const destPath = path.resolve(SOZAI_DIR, destName);

  const media = await drive.files.get(
    {
      fileId,
      alt: 'media',
    },
    {
      responseType: 'stream',
    }
  );

  await pipeline(media.data, fs.createWriteStream(destPath));
  return {
    imagePath: destPath,
    sourceKey: `drive:${fileId}:${metadata.data.modifiedTime ?? ''}`,
  };
};

const resolveImageExtFromResponse = (url, contentType = '', fallback = '.png') => {
  const ct = String(contentType ?? '').toLowerCase();
  if (ct.includes('image/png')) return '.png';
  if (ct.includes('image/jpeg') || ct.includes('image/jpg')) return '.jpg';
  if (ct.includes('image/webp')) return '.webp';
  const pathname = String(url ?? '').split('?')[0];
  const ext = path.extname(pathname).toLowerCase();
  if (isSupportedImageExt(ext)) {
    return ext;
  }
  return fallback;
};

const fetchPublicDriveImageResponse = async (fileId) => {
  const id = String(fileId ?? '').trim();
  if (!id) {
    throw new Error('Required data is missing.');
  }
  const candidates = [
    `https://drive.google.com/uc?export=download&id=${encodeURIComponent(id)}`,
    `https://drive.google.com/uc?export=view&id=${encodeURIComponent(id)}`,
    `https://lh3.googleusercontent.com/d/${encodeURIComponent(id)}=s2048`,
  ];

  let lastError = null;
  for (const url of candidates) {
    try {
      const res = await fetchWithTimeout(url, {method: 'GET', redirect: 'follow'}, 20000);
      if (!res.ok) {
        lastError = new Error(`HTTP ${res.status}`);
        continue;
      }
      const contentType = String(res.headers.get('content-type') ?? '').toLowerCase();
      if (contentType.includes('text/html')) {
        lastError = new Error('non-image response');
        continue;
      }
      if (!res.body) {
        lastError = new Error('empty response body');
        continue;
      }
      return res;
    } catch (error) {
      lastError = error;
    }
  }
  const detail = lastError ? ` / ${String(lastError.message || lastError)}` : '';
  throw new Error(
    `Failed to download public Drive image (${id}). Make the file public (Anyone with link, viewer) or use Google Login.${detail}`
  );
};

const downloadPublicDriveImage = async (fileId, storyId, prefix = '000-runtime-') => {
  await cleanSozaiFilesByPrefix(prefix);
  const res = await fetchPublicDriveImageResponse(fileId);
  const ext = resolveImageExtFromResponse(res.url, res.headers.get('content-type'), '.jpg');
  const destName = `${prefix}${storyId}${ext}`;
  const destPath = path.resolve(SOZAI_DIR, destName);
  await pipeline(Readable.fromWeb(res.body), fs.createWriteStream(destPath));
  return {
    imagePath: destPath,
    sourceKey: `drive-public:${fileId}:${res.headers.get('last-modified') ?? ''}`,
  };
};

const copyLocalImageToSozai = async (imageRef, storyId, prefix = '000-runtime-') => {
  ensureLocalImageDir();
  await cleanSozaiFilesByPrefix(prefix);
  const sourcePath = resolveLocalImagePathFromRef(imageRef);
  await ensureFileExists(sourcePath, 'Local image file');

  const stats = await fsp.stat(sourcePath);
  const ext = path.extname(sourcePath).toLowerCase();
  if (!imageMimeTypes[ext]) {
    throw new Error(`繝ｭ繝ｼ繧ｫ繝ｫ邏譚舌・蠖｢蠑上′譛ｪ蟇ｾ蠢懊〒縺・ ${ext}`);
  }

  const destName = `${prefix}${storyId}${ext}`;
  const destPath = path.resolve(SOZAI_DIR, destName);
  await fsp.copyFile(sourcePath, destPath);
  return {
    imagePath: destPath,
    sourceKey: `local:${path.basename(sourcePath)}:${stats.mtimeMs}`,
  };
};

const getStoryImageDataUrl = async (drive, story, {preview = false} = {}) => {
  if (story.imageFileId) {
    const cacheKey = preview ? `preview-image:${story.imageFileId}` : '';
    if (cacheKey) {
      const cached = previewImageCache.get(cacheKey);
      if (cached && Number(cached.expiresAt || 0) > Date.now()) {
        return {
          imageSrc: cached.imageSrc,
          sourceKey: cached.sourceKey,
        };
      }
    }

    if (isLocalImageRef(story.imageFileId)) {
      const local = await copyLocalImageToSozai(
        story.imageFileId,
        story.storyId,
        preview ? '000-preview-' : '000-runtime-'
      );
      const payload = {
        imageSrc: await toDataUrl(local.imagePath),
        sourceKey: local.sourceKey,
      };
      if (cacheKey) {
        setPreviewImageCache(cacheKey, {
          ...payload,
          expiresAt: Date.now() + PREVIEW_IMAGE_CACHE_TTL_MS,
        });
      }
      return {
        imageSrc: payload.imageSrc,
        sourceKey: payload.sourceKey,
      };
    }
    const result = drive
      ? await downloadDriveImage(
          drive,
          story.imageFileId,
          story.storyId,
          preview ? '000-preview-' : '000-runtime-'
        )
      : await downloadPublicDriveImage(story.imageFileId, story.storyId, preview ? '000-preview-' : '000-runtime-');
    const payload = {
      imageSrc: await toDataUrl(result.imagePath),
      sourceKey: result.sourceKey,
    };
    if (cacheKey) {
      setPreviewImageCache(cacheKey, {
        ...payload,
        expiresAt: Date.now() + PREVIEW_IMAGE_CACHE_TTL_MS,
      });
    }
    return {
      imageSrc: payload.imageSrc,
      sourceKey: payload.sourceKey,
    };
  }

  throw new Error('Required data is missing.');
};

const renderPreviewStill = async ({story, scene, drive}) => {
  await fsp.mkdir(OUTPUT_DIR, {recursive: true});
  const image = await getStoryImageDataUrl(drive, story, {preview: true});
  const cacheSeed = {
    v: PREVIEW_CACHE_VERSION,
    storyId: story.storyId,
    sceneNo: scene.sceneNo,
    text: scene.textRaw,
    theme: scene.theme,
    font: scene.font,
    color: scene.color,
    strokeColor: scene.strokeColor,
    strokeWidth: scene.strokeWidth,
    telopType: scene.telopType,
    textPosition: scene.textPosition,
    sourceKey: image.sourceKey,
    mode: PREVIEW_MODE,
    scale: PREVIEW_SCALE,
    quality: PREVIEW_JPEG_QUALITY,
  };
  const hash = createHash('sha1').update(JSON.stringify(cacheSeed)).digest('hex').slice(0, 18);
  const filename = `preview_${hash}.jpg`;
  const outputPath = path.resolve(OUTPUT_DIR, filename);
  const url = `/output/${filename}`;

  if (fs.existsSync(outputPath)) {
    return {filename, outputPath, url, cached: true};
  }

  return withPreviewLock(hash, async () => {
    if (fs.existsSync(outputPath)) {
      return {filename, outputPath, url, cached: true};
    }

    const serveUrl = await getBundleLocation();
    const inputProps = {
      imageSrc: image.imageSrc,
      text: scene.textRaw,
      theme: scene.theme,
      font: scene.font,
      color: scene.color,
      strokeColor: scene.strokeColor,
      strokeWidth: scene.strokeWidth,
      telopType: scene.telopType,
      textPosition: scene.textPosition,
      previewMode: PREVIEW_MODE,
    };

    const composition = await selectComposition({
      serveUrl,
      id: 'CaptionStill',
      inputProps,
    });

    await renderStill({
      composition,
      serveUrl,
      output: outputPath,
      inputProps,
      imageFormat: 'jpeg',
      jpegQuality: PREVIEW_JPEG_QUALITY,
      scale: PREVIEW_SCALE,
      overwrite: true,
    });

    return {filename, outputPath, url, cached: false};
  });
};

const runCommandWithLogs = async (command, args, cwd, onLog) =>
  new Promise((resolve, reject) => {
    const child = spawn([command, ...args].join(' '), {
      cwd,
      env: process.env,
      shell: true,
      windowsHide: true,
    });

    const flush = (chunk) => {
      const text = chunk.toString('utf8');
      text
        .split(/\r?\n/)
        .map((line) => line.trimEnd())
        .filter((line) => line.length > 0)
        .forEach((line) => onLog(line));
    };

    child.stdout.on('data', flush);
    child.stderr.on('data', flush);
    child.on('error', (error) => reject(error));
    child.on('close', (code) => {
      if (code === 0) {
        resolve();
        return;
      }
      reject(new Error(`繧ｳ繝槭Φ繝牙ｮ溯｡後↓螟ｱ謨励＠縺ｾ縺励◆ (exit=${code})`));
    });
  });

const createJobId = () => `JOB_${Date.now()}_${Math.floor(Math.random() * 1000)}`;

const createJobRecordInSheet = async (sheets, spreadsheetId, storyId, status) => {
  const _ignore = {sheets, spreadsheetId};
  void _ignore;
  const jobId = createJobId();
  const initialValues = [jobId, storyId, status, '', '', '', '', 'output'];

  return {
    jobId,
    rowNumber: null,
    values: initialValues,
  };
};

const writeJobRow = async (sheets, spreadsheetId, rowNumber, values) => {
  const _ignore = {sheets, spreadsheetId, rowNumber, values};
  void _ignore;
};

const enqueueStoryRenderJob = async ({sheets, spreadsheetId, storyId, validateStory = true}) => {
  if (validateStory) {
    await getStoryDetail(sheets, spreadsheetId, storyId);
  }

  const sheetJob = await createJobRecordInSheet(sheets, spreadsheetId, storyId, STATUS.pending);
  const job = {
    jobId: sheetJob.jobId,
    storyId,
    status: STATUS.pending,
    createdAt: nowIso(),
    startedAt: '',
    finishedAt: '',
    outputDir: 'output',
    outputFileId: '',
    error: '',
    logs: [''],
    sheetRowNumber: sheetJob.rowNumber,
    sheetValues: sheetJob.values,
  };
  jobs.set(job.jobId, job);
  queueRenderJob(job.jobId, runRenderJob);
  return job;
};

const hasActiveJobForStory = (storyId) =>
  Array.from(jobs.values()).some(
    (job) => job.storyId === storyId && (job.status === STATUS.pending || job.status === STATUS.rendering)
  );

const storyHasRenderableScenes = (scenes) => scenes.some((scene) => scene.textRaw.trim().length > 0);

const selectStoriesByInclusiveRange = (stories, startStoryId, endStoryId) => {
  const list = Array.isArray(stories) ? stories : [];
  if (list.length === 0) {
    return {
      stories: [],
      rangeStartId: '',
      rangeEndId: '',
    };
  }

  const startId = String(startStoryId ?? '').trim();
  const endId = String(endStoryId ?? '').trim();
  if (!startId && !endId) {
    return {
      stories: list,
      rangeStartId: '',
      rangeEndId: '',
    };
  }

  const storyIds = list.map((story) => String(story.storyId ?? '').trim());
  let startIndex = startId ? storyIds.indexOf(startId) : 0;
  let endIndex = endId ? storyIds.indexOf(endId) : storyIds.length - 1;

  if (startId && startIndex < 0) {
    throw new Error(`開始ストーリーIDが見つかりません: ${startId}`);
  }
  if (endId && endIndex < 0) {
    throw new Error(`終了ストーリーIDが見つかりません: ${endId}`);
  }

  if (startIndex > endIndex) {
    const temp = startIndex;
    startIndex = endIndex;
    endIndex = temp;
  }

  return {
    stories: list.slice(startIndex, endIndex + 1),
    rangeStartId: storyIds[startIndex] || '',
    rangeEndId: storyIds[endIndex] || '',
  };
};

const enqueueFromSheetStories = async ({
  sheets,
  spreadsheetId,
  limit = Number.POSITIVE_INFINITY,
  skipCompleted = false,
  startStoryId = '',
  endStoryId = '',
}) => {
  const stories = await getStories(sheets, spreadsheetId);
  const ranged = selectStoriesByInclusiveRange(stories, startStoryId, endStoryId);
  const filtered = skipCompleted
    ? ranged.stories.filter(
        (story) =>
          story.status !== STATUS.done &&
          story.status !== STATUS.error &&
          story.status !== STATUS.pending &&
          story.status !== STATUS.rendering
      )
    : ranged.stories;
  if (filtered.length === 0) {
    return {
      considered: 0,
      queuedCount: 0,
      skippedCount: 0,
      queued: [],
      skipped: [],
      rangeStartId: ranged.rangeStartId,
      rangeEndId: ranged.rangeEndId,
    };
  }
  const allScenes = await getAllScenes(sheets, spreadsheetId);
  const renderableStoryIds = new Set(
    allScenes
      .filter((scene) => scene.textRaw.trim().length > 0)
      .map((scene) => scene.storyId)
  );
  const normalizedLimit =
    Number.isFinite(limit) && limit > 0 ? Math.max(1, Number.parseInt(String(limit), 10)) : filtered.length;
  const candidates = filtered.slice(0, normalizedLimit);

  const queued = [];
  const skipped = [];
  for (const story of candidates) {
    if (hasActiveJobForStory(story.storyId)) {
      skipped.push({storyId: story.storyId, reason: 'active-job'});
      continue;
    }

    if (!renderableStoryIds.has(story.storyId)) {
      skipped.push({storyId: story.storyId, reason: 'no-renderable-scenes'});
      continue;
    }

    const job = await enqueueStoryRenderJob({
      sheets,
      spreadsheetId,
      storyId: story.storyId,
      validateStory: false,
    });
    queued.push({storyId: story.storyId, jobId: job.jobId});
  }

  return {
    considered: candidates.length,
    queuedCount: queued.length,
    skippedCount: skipped.length,
    queued,
    skipped,
    rangeStartId: ranged.rangeStartId,
    rangeEndId: ranged.rangeEndId,
  };
};

const pushJobLog = (jobId, message) => {
  const job = jobs.get(jobId);
  if (!job) return;
  const time = new Date().toLocaleTimeString('ja-JP');
  job.logs.push(`[${time}] ${message}`);
  if (job.logs.length > 200) {
    job.logs.shift();
  }
};

const setJobState = (jobId, patch) => {
  const job = jobs.get(jobId);
  if (!job) return null;
  Object.assign(job, patch);
  return job;
};

const getJobList = () =>
  Array.from(jobs.values())
    .map((job) => ({
      ...job,
      logs: [...job.logs],
    }))
    .sort((a, b) => String(b.createdAt).localeCompare(String(a.createdAt)));

const queueRenderJob = (jobId, runner) => {
  jobQueueTail = jobQueueTail
    .then(() => runner(jobId))
    .catch((error) => {
      console.error(`[render-queue] ${error.message}`);
    });
};

const buildRenderCsvRows = (storyId, scenes, renderStyle) => {
  const style = normalizeRenderStyle(renderStyle);
  const rows = scenes
    .filter((scene) => scene.textRaw.trim().length > 0)
    .sort((a, b) => a.sceneNo - b.sceneNo)
    .map((scene) => {
      const baseName = scene.outputFileName || `${storyId}_${String(scene.sceneNo).padStart(2, '0')}.png`;
      const normalizedFileName = baseName.replace(/[\\/]/g, '-');
      return [
        storyId,
        scene.sceneNo,
        normalizedFileName,
        scene.textRaw,
        STATUS.pending,
        style.theme || scene.theme,
        style.telopType || scene.telopType,
        style.textPosition || scene.textPosition,
        style.font || scene.font,
        style.color || scene.color,
        style.strokeColor || scene.strokeColor,
        style.strokeWidth ?? scene.strokeWidth,
      ];
    });

  if (rows.length === 0) {
    throw new Error('Required data is missing.');
  }

  return [
    [
      'ストーリーID',
      'シーン番号',
      '出力ファイル名',
      '本文（入力）',
      '状態',
      'テーマ',
      'テロップ種類',
      'テキスト位置',
      'フォント',
      '文字色',
      'フチ色',
      'フチ太さ',
    ],
    ...rows,
  ];
};

const markScenesAsDone = async (sheets, spreadsheetId, storyId, sceneNos) => {
  const _ignore = {sheets, spreadsheetId};
  void _ignore;
  const meta = readLocalMeta();
  const storyMeta = ensureStoryMeta(meta, storyId);
  for (const sceneNo of sceneNos) {
    const key = String(sceneNo);
    if (!storyMeta.scenes[key] || typeof storyMeta.scenes[key] !== 'object') {
      storyMeta.scenes[key] = {};
    }
    storyMeta.scenes[key].status = STATUS.done;
  }
  writeLocalMeta(meta);
};

const runRenderJob = async (jobId) => {
  const job = jobs.get(jobId);
  if (!job) return;

  const config = getRuntimeConfig();
  const spreadsheetId = config.spreadsheetId;
  const clients = await tryBuildGoogleClients(config);
  const sheets = clients.sheets;
  const drive = clients.drive;
  if (sheets) {
    await ensureSpreadsheetStructure(sheets, spreadsheetId);
  } else {
    pushJobLog(jobId, 'Running in public spreadsheet mode (no Google login).');
  }

  const setSheetJobValues = async (nextValues) => {
    job.sheetValues = nextValues;
    await writeJobRow(sheets, spreadsheetId, job.sheetRowNumber, nextValues);
  };

  let csvBackup = null;
  try {
    setJobState(jobId, {
      status: STATUS.rendering,
      startedAt: nowIso(),
      finishedAt: '',
      error: '',
    });
    pushJobLog(jobId, 'Render started.');

    await setSheetJobValues([
      job.jobId,
      job.storyId,
      STATUS.rendering,
      job.startedAt,
      '',
      '',
      '',
      job.outputDir,
    ]);

    await updateStoryStatus(sheets, spreadsheetId, job.storyId, STATUS.rendering);

    const {story, scenes} = await getStoryDetail(sheets, spreadsheetId, job.storyId);
    const csvRows = buildRenderCsvRows(job.storyId, scenes, config.renderStyle);
    const csvContent = `${buildCsv(csvRows)}\n`;
    csvBackup = fs.existsSync(CSV_PATH) ? await fsp.readFile(CSV_PATH, 'utf8') : null;

    await fsp.writeFile(CSV_PATH, csvContent, 'utf8');
    pushJobLog(jobId, `CSV繧呈峩譁ｰ縺励∪縺励◆ (${csvRows.length - 1}繧ｷ繝ｼ繝ｳ)`);
    await cleanSozaiFilesByPrefix('000-preview-');

    if (!story.imageFileId) {
      throw new Error('????ID??????????????C??Drive file ID???? local:...???????????');
    }
    if (isLocalImageRef(story.imageFileId)) {
      const local = await copyLocalImageToSozai(story.imageFileId, story.storyId, '000-runtime-');
      pushJobLog(jobId, `?????????: ${path.basename(local.imagePath)}`);
    } else {
      const downloaded = drive
        ? await downloadDriveImage(drive, story.imageFileId, story.storyId)
        : await downloadPublicDriveImage(story.imageFileId, story.storyId);
      pushJobLog(jobId, `Drive?????: ${path.basename(downloaded.imagePath)}`);
    }

    await fsp.mkdir(OUTPUT_DIR, {recursive: true});
    await runCommandWithLogs(NPM_COMMAND, ['run', 'render:csv'], PROJECT_ROOT, (line) => {
      pushJobLog(jobId, line);
    });

    const renderedSceneNos = new Set(
      scenes
        .filter((scene) => scene.textRaw.trim().length > 0)
        .map((scene) => scene.sceneNo)
    );

    await markScenesAsDone(sheets, spreadsheetId, story.storyId, renderedSceneNos);
    await updateStoryStatus(sheets, spreadsheetId, story.storyId, STATUS.done);

    setJobState(jobId, {
      status: STATUS.done,
      finishedAt: nowIso(),
      outputFileId: '',
      error: '',
    });
    pushJobLog(jobId, 'Render completed.');

    await setSheetJobValues([
      job.jobId,
      job.storyId,
      STATUS.done,
      job.startedAt,
      job.finishedAt,
      '',
      '',
      job.outputDir,
    ]);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    setJobState(jobId, {
      status: STATUS.error,
      finishedAt: nowIso(),
      error: message,
    });
    pushJobLog(jobId, `繧ｨ繝ｩ繝ｼ: ${message}`);

    try {
      await updateStoryStatus(sheets, spreadsheetId, job.storyId, STATUS.error);
      await writeJobRow(sheets, spreadsheetId, job.sheetRowNumber, [
        job.jobId,
        job.storyId,
        STATUS.error,
        job.startedAt || '',
        job.finishedAt || '',
        '',
        message,
        job.outputDir,
      ]);
    } catch (updateError) {
      console.error(`[render-job] failed to persist error status: ${String(updateError)}`);
    }
  } finally {
    if (csvBackup !== null) {
      await fsp.writeFile(CSV_PATH, csvBackup, 'utf8');
    }
  }
};

const app = express();
app.use(express.json({limit: '25mb'}));
app.use(express.static(PUBLIC_DIR));
app.use('/output', express.static(OUTPUT_DIR));

app.get('/api/health', (_req, res) => {
  res.json({ok: true, now: nowIso()});
});

app.get(
  '/api/auth/google/status',
  asyncHandler(async (_req, res) => {
    const config = getRuntimeConfig();
    const info = resolveOAuthClientInfo(config);
    res.json({
      ok: true,
      ready: info.ready,
      hasTokenFile: Boolean(info.tokenJson),
      hasRefreshToken: info.hasRefreshToken,
      tokenPath: info.tokenPath,
      redirectUri: getOAuthRedirectUri(),
    });
  })
);

app.post(
  '/api/auth/google/start',
  asyncHandler(async (_req, res) => {
    const config = getRuntimeConfig();
    const info = resolveOAuthClientInfo(config);
    if (!info.ready) {
      throw new Error(
        'OAuth client info was not found. Set GOOGLE_CLIENT_ID/GOOGLE_CLIENT_SECRET or use an existing token file.'
      );
    }

    pruneOAuthAuthStates();
    const state = randomBytes(24).toString('hex');
    oauthAuthStates.set(state, {
      createdAt: Date.now(),
      tokenPath: info.tokenPath,
      clientId: info.clientId,
      clientSecret: info.clientSecret,
    });

    const redirectUri = getOAuthRedirectUri();
    const oauth2Client = buildOAuthClient({
      clientId: info.clientId,
      clientSecret: info.clientSecret,
      redirectUri,
    });
    const authUrl = oauth2Client.generateAuthUrl({
      access_type: 'offline',
      prompt: 'consent',
      include_granted_scopes: true,
      scope: GOOGLE_OAUTH_SCOPES,
      state,
    });

    res.json({
      ok: true,
      authUrl,
      redirectUri,
      tokenPath: info.tokenPath,
    });
  })
);

app.get(
  '/api/auth/google/callback',
  asyncHandler(async (req, res) => {
    const code = String(req.query.code ?? '').trim();
    const state = String(req.query.state ?? '').trim();
    if (!code || !state) {
      res
        .status(400)
        .send(
          renderOAuthPopupPage({
            ok: false,
            message: '',
            payload: {type: 'google-auth-complete', ok: false},
          })
        );
      return;
    }

    pruneOAuthAuthStates();
    const entry = oauthAuthStates.get(state);
    if (!entry) {
      res
        .status(400)
        .send(
          renderOAuthPopupPage({
            ok: false,
            message: '',
            payload: {type: 'google-auth-complete', ok: false},
          })
        );
      return;
    }
    oauthAuthStates.delete(state);

    try {
      const redirectUri = getOAuthRedirectUri();
      const oauth2Client = buildOAuthClient({
        clientId: entry.clientId,
        clientSecret: entry.clientSecret,
        redirectUri,
      });
      const tokenResponse = await oauth2Client.getToken(code);
      const saved = await saveOAuthToken({
        tokenPath: entry.tokenPath,
        clientId: entry.clientId,
        clientSecret: entry.clientSecret,
        redirectUri,
        tokens: tokenResponse.tokens || {},
      });
      const hasRefreshToken = Boolean(String(saved.refresh_token ?? '').trim());

      res.send(
        renderOAuthPopupPage({
          ok: true,
          message: hasRefreshToken
            ? 'OAuth token saved. Return to GUI and reload.'
            : '',
          payload: {
            type: 'google-auth-complete',
            ok: hasRefreshToken,
            hasRefreshToken,
            tokenPath: entry.tokenPath,
          },
        })
      );
    } catch (error) {
      res
        .status(500)
        .send(
          renderOAuthPopupPage({
            ok: false,
            message: error instanceof Error ? error.message : String(error),
            payload: {type: 'google-auth-complete', ok: false},
          })
        );
    }
  })
);

app.get(
  '/api/config',
  asyncHandler(async (_req, res) => {
    res.json({
      config: getRuntimeConfig(),
      defaultTokenPath: DEFAULT_TOKEN_PATH,
    });
  })
);

app.put(
  '/api/config',
  asyncHandler(async (req, res) => {
    const body = req.body ?? {};
    const patch = {};

    if (Object.prototype.hasOwnProperty.call(body, 'spreadsheetId') || Object.prototype.hasOwnProperty.call(body, 'spreadsheetUrl')) {
      const spreadsheetIdInput = String(body.spreadsheetId ?? body.spreadsheetUrl ?? '').trim();
      patch.spreadsheetId = parseSpreadsheetIdInput(spreadsheetIdInput);
    }
    if (Object.prototype.hasOwnProperty.call(body, 'tokenPath')) {
      patch.tokenPath = String(body.tokenPath ?? '').trim();
    }
    if (Object.prototype.hasOwnProperty.call(body, 'outputDriveFolderId')) {
      patch.outputDriveFolderId = String(body.outputDriveFolderId ?? '').trim();
    }

    const hasRenderStylePayload =
      Object.prototype.hasOwnProperty.call(body, 'renderStyle') ||
      Object.prototype.hasOwnProperty.call(body, 'theme') ||
      Object.prototype.hasOwnProperty.call(body, 'telopType') ||
      Object.prototype.hasOwnProperty.call(body, 'textPosition') ||
      Object.prototype.hasOwnProperty.call(body, 'font') ||
      Object.prototype.hasOwnProperty.call(body, 'color') ||
      Object.prototype.hasOwnProperty.call(body, 'strokeColor') ||
      Object.prototype.hasOwnProperty.call(body, 'strokeWidth');

    if (hasRenderStylePayload) {
      const rawRenderStyle =
        body.renderStyle && typeof body.renderStyle === 'object'
          ? body.renderStyle
          : {
              theme: body.theme,
              telopType: body.telopType,
              textPosition: body.textPosition,
              font: body.font,
              color: body.color,
              strokeColor: body.strokeColor,
              strokeWidth: body.strokeWidth,
            };
      patch.renderStyle = normalizeRenderStyle(rawRenderStyle);
    }

    const stored = writeLocalConfig(patch);
    res.json({
      ok: true,
      config: {
        spreadsheetId: stored.spreadsheetId || '',
        tokenPath: stored.tokenPath || '',
        outputDriveFolderId: stored.outputDriveFolderId || '',
        renderStyle: normalizeRenderStyle(stored.renderStyle),
      },
    });
  })
);

app.post(
  '/api/init',
  asyncHandler(async (_req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }
    const {sheets} = await buildGoogleClients(config);
    await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    res.json({ok: true});
  })
);

app.get(
  '/api/bootstrap',
  asyncHandler(async (_req, res) => {
    const config = getRuntimeConfig();
    const authInfo = resolveOAuthClientInfo(config);
    const response = {
      config,
      statuses: STATUS_LIST,
      telopTypes: TELOP_TYPES,
      textPositions: TEXT_POSITIONS,
      themes: [],
      stories: [],
      jobs: getJobList(),
      needsAuth: false,
      auth: {
        ready: authInfo.ready,
        hasTokenFile: Boolean(authInfo.tokenJson),
        hasRefreshToken: authInfo.hasRefreshToken,
        tokenPath: authInfo.tokenPath,
        redirectUri: getOAuthRedirectUri(),
        message: '',
      },
    };

    if (!config.spreadsheetId) {
      res.json(response);
      return;
    }
    try {
      const clients = await tryBuildGoogleClients(config);
      const sheets = clients.sheets;
      if (sheets) {
        await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
      }
      response.themes = await getThemes(sheets, config.spreadsheetId);
      response.stories = await getStories(sheets, config.spreadsheetId);
      response.needsAuth = false;
      if (!sheets) {
        response.auth.message = 'Public spreadsheet read mode is active.';
      }
    } catch (_error) {
      response.needsAuth = true;
      response.auth.message =
        'Could not read public spreadsheet. Share it as Anyone with the link (Viewer) or run Google Login.';
    }
    res.json(response);
  })
);

app.get(
  '/api/sheets/guide',
  asyncHandler(async (_req, res) => {
    const config = getRuntimeConfig();
    const payload = buildSheetGuidePayload(config);
    res.json({
      ok: true,
      ...payload,
      tips: [
        `1) ストーリー一覧A列に最大${BATCH_MAX_STORIES}件まで入力できます。`,
        `2) シーン一覧A-C列には最大${BATCH_MAX_STORIES * SCENES_PER_STORY}行まで入力できます。`,
        '3) スプレッドシートは本文テキスト中心で運用してください。',
        '4) 詳細な見た目設定はGUIで変更してください。',
      ],
    });
  })
);

app.post(
  '/api/stories/:storyId/preview-image',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    const storyId = String(req.params.storyId ?? '').trim();
    const directImageFileId = String(req.body?.imageFileId ?? '').trim();
    const knownSourceKey = String(req.body?.knownSourceKey ?? '').trim();
    const clients = await tryBuildGoogleClients(config);
    const sheets = clients.sheets;
    const drive = clients.drive;

    let story = null;
    if (directImageFileId) {
      story = {storyId, imageFileId: directImageFileId};
    } else {
      if (!config.spreadsheetId) {
        throw new Error('Required data is missing.');
      }
      if (sheets) {
        await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
      }
      const detail = await getStoryDetail(sheets, config.spreadsheetId, storyId);
      story = detail.story;
    }

    const image = await getStoryImageDataUrl(drive, story, {preview: true});
    if (knownSourceKey && knownSourceKey === image.sourceKey) {
      res.json({
        ok: true,
        unchanged: true,
        sourceKey: image.sourceKey,
      });
      return;
    }

    res.json({
      ok: true,
      unchanged: false,
      sourceKey: image.sourceKey,
      imageSrc: image.imageSrc,
    });
  })
);

app.post(
  '/api/stories/:storyId/preview',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    const storyId = String(req.params.storyId ?? '').trim();
    const scenePayload = req.body?.scene ?? {};
    const directImageFileId = String(req.body?.imageFileId ?? '').trim();
    const clients = await tryBuildGoogleClients(config);
    const sheets = clients.sheets;
    const drive = clients.drive;

    let story = null;
    let baseScene = null;

    if (directImageFileId) {
      story = {storyId, imageFileId: directImageFileId};
    } else {
      if (!config.spreadsheetId) {
        throw new Error('Required data is missing.');
      }
      if (sheets) {
        await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
      }
      const detail = await getStoryDetail(sheets, config.spreadsheetId, storyId);
      story = detail.story;
      baseScene = detail.scenes.find((scene) => scene.sceneNo === Number(scenePayload.sceneNo)) || detail.scenes[0];
    }

    const mergedScene = normalizeScenePayload(storyId, {
      ...(baseScene || {}),
      ...normalizeRenderStyle(config.renderStyle),
      ...scenePayload,
    });
    const preview = await renderPreviewStill({
      story,
      scene: mergedScene,
      drive,
    });
    res.json({
      ok: true,
      previewUrl: preview.url,
      cached: preview.cached,
      previewMode: PREVIEW_MODE,
      sceneNo: mergedScene.sceneNo,
      telopType: mergedScene.telopType,
      textPosition: mergedScene.textPosition,
    });
  })
);

app.post(
  '/api/local-image',
  asyncHandler(async (req, res) => {
    const fileNameInput = String(req.body?.filename ?? '').trim();
    const dataUrl = String(req.body?.dataUrl ?? '').trim();
    if (!dataUrl) {
      throw new Error('Required data is missing.');
    }

    ensureAppDataDir();
    ensureLocalImageDir();
    const parsed = parseImageDataUrl(dataUrl);
    const maxBytes = 12 * 1024 * 1024;
    if (parsed.buffer.length > maxBytes) {
      throw new Error('Required data is missing.');
    }

    const ext = path.extname(fileNameInput).toLowerCase() || extFromUploadedMime(parsed.mimeType);
    if (!imageMimeTypes[ext]) {
      throw new Error('Unsupported image extension: ' + ext);
    }

    const stem = sanitizeLocalFileName(path.basename(fileNameInput, path.extname(fileNameInput))) || 'local_image';
    const fileName = `${stem}_${Date.now()}_${Math.floor(Math.random() * 10000)}${ext}`;
    const outputPath = path.resolve(LOCAL_IMAGE_DIR, fileName);
    await fsp.writeFile(outputPath, parsed.buffer);

    res.json({
      ok: true,
      imageRef: `${LOCAL_IMAGE_REF_PREFIX}${fileName}`,
      fileName,
    });
  })
);

app.get(
  '/api/stories',
  asyncHandler(async (_req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }
    const clients = await tryBuildGoogleClients(config);
    const sheets = clients.sheets;
    if (sheets) {
      await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    }
    const stories = await getStories(sheets, config.spreadsheetId);
    res.json({stories, accessMode: sheets ? 'auth' : 'public'});
  })
);

app.put(
  '/api/stories/:storyId/image',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }
    const storyId = String(req.params.storyId ?? '').trim();
    const imageRef = String(req.body?.imageRef ?? '').trim();
    const {sheets} = await buildGoogleClients(config);
    await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    await updateStoryImageRef(sheets, config.spreadsheetId, storyId, imageRef);
    res.json({ok: true, storyId, imageRef});
  })
);

app.post(
  '/api/stories',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }
    const {sheets} = await buildGoogleClients(config);
    await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    const story = await createStory(sheets, config.spreadsheetId, req.body ?? {});
    res.json({ok: true, story});
  })
);

app.post(
  '/api/batch/create-and-render',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }

    const imageRefs = Array.isArray(req.body?.imageRefs)
      ? req.body.imageRefs.map((value) => String(value ?? '').trim()).filter((value) => value.length > 0)
      : [];
    if (imageRefs.length === 0) {
      throw new Error('Required data is missing.');
    }
    if (imageRefs.length > BATCH_MAX_STORIES) {
      throw new Error(`Batch size must be <= ${BATCH_MAX_STORIES} stories.`);
    }

    const sceneTexts = Array.isArray(req.body?.sceneTexts)
      ? req.body.sceneTexts.map((value) => String(value ?? ''))
      : [];
    if (sceneTexts.length !== SCENES_PER_STORY) {
      throw new Error(`sceneTexts must contain exactly ${SCENES_PER_STORY} lines.`);
    }

    const theme = String(req.body?.theme ?? '').trim() || 'clean';
    const titlePrefix = String(req.body?.titlePrefix ?? '').trim() || '荳諡ｬ繧ｹ繝医・繝ｪ繝ｼ';
    const telopType = String(req.body?.telopType ?? '').trim() || TELOP_TYPES[0] || '標準';
    const textPosition = String(req.body?.textPosition ?? '').trim() || TEXT_POSITIONS[1] || '中央';
    const font = String(req.body?.font ?? '').trim() || '標準';
    const color = normalizeHexColor(req.body?.color, '#FFFFFF');
    const strokeColor = normalizeHexColor(req.body?.strokeColor, '#121212');
    const strokeWidth = normalizeStrokeWidth(req.body?.strokeWidth, 0.8);
    const enqueueRender = req.body?.enqueueRender !== false;

    const {sheets} = await buildGoogleClients(config);
    await ensureSpreadsheetStructure(sheets, config.spreadsheetId);

    const created = [];
    for (let i = 0; i < imageRefs.length; i += 1) {
      const index = i + 1;
      const title = `${titlePrefix}_${String(index).padStart(2, '0')}`;
      const story = await createStory(sheets, config.spreadsheetId, {
        title,
        theme,
        imageFileId: imageRefs[i],
      });

      const scenes = buildScenePayloadsFromTemplate({
        storyId: story.storyId,
        theme,
        sceneTexts,
        telopType,
        textPosition,
        font,
        color,
        strokeColor,
        strokeWidth,
      });
      await saveScenes(sheets, config.spreadsheetId, story.storyId, scenes);

      let jobId = '';
      if (enqueueRender) {
        const job = await enqueueStoryRenderJob({
          sheets,
          spreadsheetId: config.spreadsheetId,
          storyId: story.storyId,
        });
        jobId = job.jobId;
      }

      created.push({
        index,
        storyId: story.storyId,
        title,
        imageRef: imageRefs[i],
        jobId,
      });
    }

    res.json({
      ok: true,
      createdCount: created.length,
      estimatedImageCount: created.length * sceneTexts.length,
      enqueueRender,
      created,
    });
  })
);

app.post(
  '/api/batch/enqueue-from-sheet',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }
    const limitRaw = Number.parseInt(String(req.body?.limit ?? String(BATCH_MAX_STORIES)), 10);
    const limit = Number.isFinite(limitRaw)
      ? Math.min(Math.max(limitRaw, 1), BATCH_MAX_STORIES)
      : BATCH_MAX_STORIES;

    const clients = await tryBuildGoogleClients(config);
    const sheets = clients.sheets;
    if (sheets) {
      await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    }
    const result = await enqueueFromSheetStories({
      sheets,
      spreadsheetId: config.spreadsheetId,
      limit,
      skipCompleted: false,
    });

    res.json({
      ok: true,
      limit,
      accessMode: sheets ? 'auth' : 'public',
      ...result,
    });
  })
);

app.post(
  '/api/batch/enqueue-all-from-sheet',
  asyncHandler(async (_req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Spreadsheet ID is not configured.');
    }
    const clients = await tryBuildGoogleClients(config);
    const sheets = clients.sheets;
    if (sheets) {
      await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    }
    const result = await enqueueFromSheetStories({
      sheets,
      spreadsheetId: config.spreadsheetId,
      limit: Number.POSITIVE_INFINITY,
      skipCompleted: true,
    });
    res.json({
      ok: true,
      mode: 'all-ready',
      accessMode: sheets ? 'auth' : 'public',
      ...result,
    });
  })
);

app.post(
  '/api/batch/enqueue-range-from-sheet',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Spreadsheet ID is not configured.');
    }

    const startStoryId = String(req.body?.startStoryId ?? '').trim();
    const endStoryId = String(req.body?.endStoryId ?? '').trim();
    if (!startStoryId && !endStoryId) {
      throw new Error('開始IDか終了IDを入力してください。');
    }

    const clients = await tryBuildGoogleClients(config);
    const sheets = clients.sheets;
    if (sheets) {
      await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    }
    const result = await enqueueFromSheetStories({
      sheets,
      spreadsheetId: config.spreadsheetId,
      limit: Number.POSITIVE_INFINITY,
      skipCompleted: true,
      startStoryId,
      endStoryId,
    });
    res.json({
      ok: true,
      mode: 'range-ready',
      accessMode: sheets ? 'auth' : 'public',
      ...result,
    });
  })
);

app.get(
  '/api/stories/:storyId/scenes',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }
    const storyId = String(req.params.storyId ?? '').trim();
    const clients = await tryBuildGoogleClients(config);
    const sheets = clients.sheets;
    if (sheets) {
      await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    }
    const detail = await getStoryDetail(sheets, config.spreadsheetId, storyId);
    res.json({
      ...detail,
      accessMode: sheets ? 'auth' : 'public',
    });
  })
);

app.put(
  '/api/stories/:storyId/scenes',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }
    const storyId = String(req.params.storyId ?? '').trim();
    const scenes = Array.isArray(req.body?.scenes) ? req.body.scenes : [];
    if (scenes.length === 0) {
      throw new Error('Required data is missing.');
    }
    const {sheets} = await buildGoogleClients(config);
    await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    const result = await saveScenes(sheets, config.spreadsheetId, storyId, scenes);
    res.json({ok: true, ...result});
  })
);

app.post(
  '/api/stories/:storyId/render',
  asyncHandler(async (req, res) => {
    const config = getRuntimeConfig();
    if (!config.spreadsheetId) {
      throw new Error('Required data is missing.');
    }
    const storyId = String(req.params.storyId ?? '').trim();
    const clients = await tryBuildGoogleClients(config);
    const sheets = clients.sheets;
    if (sheets) {
      await ensureSpreadsheetStructure(sheets, config.spreadsheetId);
    }
    const job = await enqueueStoryRenderJob({
      sheets,
      spreadsheetId: config.spreadsheetId,
      storyId,
    });

    res.json({
      ok: true,
      jobId: job.jobId,
      status: job.status,
      accessMode: sheets ? 'auth' : 'public',
    });
  })
);

app.get(
  '/api/jobs',
  asyncHandler(async (_req, res) => {
    res.json({jobs: getJobList()});
  })
);

app.get(
  '/api/jobs/:jobId',
  asyncHandler(async (req, res) => {
    const jobId = String(req.params.jobId ?? '').trim();
    const job = jobs.get(jobId);
    if (!job) {
      res.status(404).json({error: 'Job not found.'});
      return;
    }
    res.json({...job, logs: [...job.logs]});
  })
);

app.get('*', (_req, res) => {
  res.sendFile(path.resolve(PUBLIC_DIR, 'index.html'));
});

app.use((error, _req, res, _next) => {
  const message = error instanceof Error ? error.message : String(error);
  res.status(400).json({error: message});
});

app.listen(PORT, () => {
  console.log(`Local GUI started: http://localhost:${PORT}`);
});



