import path from 'node:path';
import fs from 'node:fs';
import {bundle} from '@remotion/bundler';
import {renderStill, selectComposition} from '@remotion/renderer';

const projectRoot = process.cwd();
const csvPath = path.resolve(projectRoot, 'templates/captions.csv');
const outputDirInput = String(process.env.APP_OUTPUT_DIR ?? 'output').trim() || 'output';
const outputDir = path.isAbsolute(outputDirInput)
  ? path.resolve(outputDirInput)
  : path.resolve(projectRoot, outputDirInput);
const sozaiDir = path.resolve(projectRoot, 'sozai');
const entryPoint = path.resolve(projectRoot, 'src/index.ts');

const imageMimeTypes = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
};

const statuses = {
  draft: '下書き',
  pending: '待機中',
  rendering: '書き出し中',
  done: '完了',
  error: 'エラー',
};

const renderableStatuses = new Set([statuses.pending, statuses.rendering]);

const statusAliases = new Map([
  ['下書き', statuses.draft],
  ['draft', statuses.draft],
  ['待機中', statuses.pending],
  ['pending', statuses.pending],
  ['書き出し中', statuses.rendering],
  ['rendering', statuses.rendering],
  ['完了', statuses.done],
  ['done', statuses.done],
  ['completed', statuses.done],
  ['エラー', statuses.error],
  ['error', statuses.error],
  ['failed', statuses.error],
]);

const headerAliases = {
  filename: ['filename', '出力ファイル名', 'ファイル名'],
  text: ['text', '本文', '本文入力', '本文（入力）', 'テキスト'],
  storyId: ['story_id', 'storyid', 'ストーリーid', 'ストーリーID'],
  sceneNo: ['scene_no', 'sceneno', 'シーン番号'],
  status: ['status', '状態'],
  theme: ['theme', 'テーマ'],
  telopType: ['teloptype', 'テロップ種類'],
  textPosition: ['textposition', 'テキスト位置', '文字位置'],
  font: ['font', 'フォント'],
  color: ['color', '文字色'],
  strokeColor: ['strokecolor', 'フチ色', '縁色'],
  strokeWidth: ['strokewidth', 'フチ太さ', '縁太さ'],
};

const parseCsvLine = (line) => {
  const result = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i += 1) {
    const ch = line[i];

    if (ch === '"') {
      const next = line[i + 1];
      if (inQuotes && next === '"') {
        current += '"';
        i += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (ch === ',' && !inQuotes) {
      result.push(current.trim());
      current = '';
      continue;
    }

    current += ch;
  }

  result.push(current.trim());
  return result;
};

const normalizeKey = (value) =>
  String(value ?? '')
    .trim()
    .toLowerCase()
    .replace(/[()（）]/g, '')
    .replace(/[\s_-]/g, '');

const sanitizeFilename = (value) => {
  const trimmed = String(value ?? '').trim();
  const safe = trimmed.replace(/[<>:"/\\|?*]/g, '-');
  if (!safe) {
    return '';
  }
  const ext = path.extname(safe);
  return ext ? safe : `${safe}.png`;
};

const toTwoDigit = (value, fallback) => {
  const numeric = Number.parseInt(String(value ?? ''), 10);
  if (Number.isFinite(numeric) && numeric > 0) {
    return String(numeric).padStart(2, '0');
  }
  return String(fallback).padStart(2, '0');
};

const buildDefaultFilename = (storyId, sceneNo, rowIndex) => {
  const safeStoryId = String(storyId ?? '')
    .trim()
    .replace(/[<>:"/\\|?*]/g, '-')
    .replace(/\s+/g, '-');
  const prefix = safeStoryId || 'story';
  return `${prefix}_${toTwoDigit(sceneNo, rowIndex + 1)}.png`;
};

const normalizeStatus = (value) => {
  const raw = String(value ?? '').trim();
  if (!raw) {
    return statuses.pending;
  }
  const normalized = statusAliases.get(raw.toLowerCase()) ?? statusAliases.get(raw);
  if (!normalized) {
    const allowed = Array.from(new Set(statusAliases.values())).join('/');
    throw new Error(`Unknown status "${raw}". Allowed: ${allowed}`);
  }
  return normalized;
};

const parseCsv = (content) => {
  const lines = content
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (lines.length < 2) {
    throw new Error('templates/captions.csv must include header and at least one row.');
  }

  const [headerLine, ...dataLines] = lines;
  const headers = parseCsvLine(headerLine).map((header) => header.trim());
  const normalizedHeaders = headers.map((header) => normalizeKey(header));

  const getHeaderIndex = (aliases) => {
    const aliasSet = new Set(aliases.map((alias) => normalizeKey(alias)));
    return normalizedHeaders.findIndex((header) => aliasSet.has(header));
  };

  const filenameIdx = getHeaderIndex(headerAliases.filename);
  const textIdx = getHeaderIndex(headerAliases.text);
  const storyIdIdx = getHeaderIndex(headerAliases.storyId);
  const sceneNoIdx = getHeaderIndex(headerAliases.sceneNo);
  const statusIdx = getHeaderIndex(headerAliases.status);
  const themeIdx = getHeaderIndex(headerAliases.theme);
  const telopTypeIdx = getHeaderIndex(headerAliases.telopType);
  const textPositionIdx = getHeaderIndex(headerAliases.textPosition);
  const fontIdx = getHeaderIndex(headerAliases.font);
  const colorIdx = getHeaderIndex(headerAliases.color);
  const strokeColorIdx = getHeaderIndex(headerAliases.strokeColor);
  const strokeWidthIdx = getHeaderIndex(headerAliases.strokeWidth);

  if (textIdx === -1) {
    throw new Error(
      'CSV header must include one of: text, 本文, 本文（入力）, テキスト'
    );
  }

  return dataLines.map((line, index) => {
    const cols = parseCsvLine(line);
    const text = cols[textIdx];
    const storyId = storyIdIdx >= 0 ? cols[storyIdIdx] : '';
    const sceneNo = sceneNoIdx >= 0 ? cols[sceneNoIdx] : '';
    const status = normalizeStatus(statusIdx >= 0 ? cols[statusIdx] : '');
    const theme = themeIdx >= 0 ? cols[themeIdx] : '';
    const telopType = telopTypeIdx >= 0 ? cols[telopTypeIdx] : '';
    const textPosition = textPositionIdx >= 0 ? cols[textPositionIdx] : '';
    const font = fontIdx >= 0 ? cols[fontIdx] : '';
    const color = colorIdx >= 0 ? cols[colorIdx] : '';
    const strokeColor = strokeColorIdx >= 0 ? cols[strokeColorIdx] : '';
    const strokeWidth = strokeWidthIdx >= 0 ? cols[strokeWidthIdx] : '';

    const explicitFilename =
      filenameIdx >= 0 ? sanitizeFilename(cols[filenameIdx]) : '';
    const filename =
      explicitFilename || buildDefaultFilename(storyId, sceneNo, index);

    if (!text) {
      throw new Error(`Invalid CSV row at line ${index + 2}: text is required.`);
    }

    return {
      filename,
      text,
      storyId,
      sceneNo,
      status,
      theme,
      telopType,
      textPosition,
      font,
      color,
      strokeColor,
      strokeWidth,
    };
  });
};

const findFirstImageFile = (dir) => {
  const imageEntries = fs
    .readdirSync(dir)
    .filter((name) => {
      const ext = path.extname(name).toLowerCase();
      return imageMimeTypes[ext];
    })
    .sort((a, b) => a.localeCompare(b, 'ja'));

  // Prefer the per-job runtime image downloaded by the server.
  const runtimeEntries = imageEntries.filter((name) => name.startsWith('000-runtime-'));
  const candidates = imageEntries.filter((name) => !name.startsWith('000-preview-'));
  const targetList =
    runtimeEntries.length > 0 ? runtimeEntries : candidates.length > 0 ? candidates : imageEntries;

  if (targetList.length === 0) {
    throw new Error('No image file found in sozai folder.');
  }

  return path.resolve(dir, targetList[0]);
};

const toDataUrl = (imagePath) => {
  const ext = path.extname(imagePath).toLowerCase();
  const mimeType = imageMimeTypes[ext];

  if (!mimeType) {
    throw new Error(`Unsupported image type: ${ext}`);
  }

  const buffer = fs.readFileSync(imagePath);
  return `data:${mimeType};base64,${buffer.toString('base64')}`;
};

const run = async () => {
  const csvRaw = fs.readFileSync(csvPath, 'utf8').replace(/^\uFEFF/, '');
  const rows = parseCsv(csvRaw);
  const rowsToRender = rows.filter((row) => renderableStatuses.has(row.status));
  const skipped = rows.filter((row) => !renderableStatuses.has(row.status));

  if (rowsToRender.length === 0) {
    throw new Error(
      `No renderable rows found. Use status "${statuses.pending}" or "${statuses.rendering}".`
    );
  }

  const imagePath = findFirstImageFile(sozaiDir);
  console.log(`Using source image: ${path.relative(projectRoot, imagePath)}`);
  const imageDataUrl = toDataUrl(imagePath);

  fs.mkdirSync(outputDir, {recursive: true});

  const bundleLocation = await bundle({
    entryPoint,
    webpackOverride: (config) => config,
  });

  if (skipped.length > 0) {
    const summary = skipped.reduce((acc, row) => {
      acc[row.status] = (acc[row.status] ?? 0) + 1;
      return acc;
    }, {});
    console.log(`Skipped rows by status: ${JSON.stringify(summary)}`);
  }

  for (const row of rowsToRender) {
    const output = path.resolve(outputDir, row.filename);
    fs.mkdirSync(path.dirname(output), {recursive: true});
    const composition = await selectComposition({
      serveUrl: bundleLocation,
      id: 'CaptionStill',
      inputProps: {
        imageSrc: imageDataUrl,
        text: row.text,
        theme: row.theme,
        font: row.font,
        color: row.color,
        strokeColor: row.strokeColor,
        strokeWidth: row.strokeWidth,
        telopType: row.telopType,
        textPosition: row.textPosition,
      },
    });

    await renderStill({
      composition,
      serveUrl: bundleLocation,
      output,
      inputProps: {
        imageSrc: imageDataUrl,
        text: row.text,
        theme: row.theme,
        font: row.font,
        color: row.color,
        strokeColor: row.strokeColor,
        strokeWidth: row.strokeWidth,
        telopType: row.telopType,
        textPosition: row.textPosition,
      },
      imageFormat: 'png',
      overwrite: true,
    });

    console.log(
      `Rendered: ${path.relative(projectRoot, output)} [${row.storyId || '-'} / ${
        row.sceneNo || '-'
      } / ${row.status}]`
    );
  }
};

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
