import path from 'node:path';
import fs from 'node:fs';
import {bundle} from '@remotion/bundler';
import {renderStill, selectComposition} from '@remotion/renderer';

const projectRoot = process.cwd();
const csvPath = path.resolve(projectRoot, 'templates/captions.csv');
const outputDir = path.resolve(projectRoot, 'output');
const sozaiDir = path.resolve(projectRoot, 'sozai');
const entryPoint = path.resolve(projectRoot, 'src/index.ts');

const imageMimeTypes = {
  '.png': 'image/png',
  '.jpg': 'image/jpeg',
  '.jpeg': 'image/jpeg',
  '.webp': 'image/webp',
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

const parseCsv = (content) => {
  const lines = content
    .split(/\r?\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 0);

  if (lines.length < 2) {
    throw new Error('templates/captions.csv must include header and at least one row.');
  }

  const [headerLine, ...dataLines] = lines;
  const headers = parseCsvLine(headerLine);
  const filenameIdx = headers.indexOf('filename');
  const textIdx = headers.indexOf('text');

  if (filenameIdx === -1 || textIdx === -1) {
    throw new Error('CSV header must include filename,text');
  }

  return dataLines.map((line, index) => {
    const cols = parseCsvLine(line);
    const filename = cols[filenameIdx];
    const text = cols[textIdx];

    if (!filename || !text) {
      throw new Error(`Invalid CSV row at line ${index + 2}: filename and text are required.`);
    }

    return {filename, text};
  });
};

const findFirstImageFile = (dir) => {
  const candidates = fs
    .readdirSync(dir)
    .filter((name) => {
      const ext = path.extname(name).toLowerCase();
      return imageMimeTypes[ext];
    })
    .sort((a, b) => a.localeCompare(b, 'ja'));

  if (candidates.length === 0) {
    throw new Error('No image file found in sozai folder.');
  }

  return path.resolve(dir, candidates[0]);
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

  const imagePath = findFirstImageFile(sozaiDir);
  const imageDataUrl = toDataUrl(imagePath);

  fs.mkdirSync(outputDir, {recursive: true});

  const bundleLocation = await bundle({
    entryPoint,
    webpackOverride: (config) => config,
  });

  for (const row of rows) {
    const output = path.resolve(outputDir, row.filename);
    const composition = await selectComposition({
      serveUrl: bundleLocation,
      id: 'CaptionStill',
      inputProps: {
        imageSrc: imageDataUrl,
        text: row.text,
      },
    });

    await renderStill({
      composition,
      serveUrl: bundleLocation,
      output,
      inputProps: {
        imageSrc: imageDataUrl,
        text: row.text,
      },
      imageFormat: 'png',
      overwrite: true,
    });

    console.log(`Rendered: ${path.relative(projectRoot, output)}`);
  }
};

run().catch((error) => {
  console.error(error);
  process.exit(1);
});
