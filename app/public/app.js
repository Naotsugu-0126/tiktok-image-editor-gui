const FONT_OPTIONS = ['標準', '丸ゴシック', '明朝', '太ゴシック', '角ゴシック', 'UDゴシック', '手書き'];
const STROKE_WIDTH_OPTIONS = ['0', '0.4', '0.8', '1.2', '1.6', '2.0', '2.4', '3.0'];
const PREVIEW_FALLBACK_TEXT = 'ここにプレビュー本文が表示されます。';

const state = {
  config: {renderStyle: {}},
  stories: [],
  jobs: [],
  themes: [],
  telopTypes: ['標準', '帯', '吹き出し', '角丸ラベル', '強調', 'ネオン', 'ガラス', 'アウトライン'],
  textPositions: ['上', '中央', '下'],
  auth: {needsAuth: false, message: ''},
  sceneCache: new Map(),
  previewSourceKey: '',
  previewStoryId: '',
  previewInFlight: false,
  previewPending: false,
};

const els = {
  message: document.getElementById('message'),
  sheetLocation: document.getElementById('sheet-location'),
  openSheetTop: document.getElementById('open-sheet'),
  jobsList: document.getElementById('jobs-list'),
  jobsEmpty: document.getElementById('jobs-empty'),
  spreadsheetIdInput: document.getElementById('spreadsheet-id'),
};

const ui = {
  panel: null,
  startAll: null,
  startRange: null,
  googleAuth: null,
  openSheet: null,
  reload: null,
  rangeStartId: null,
  rangeEndId: null,
  outputDir: null,
  browseOutputDir: null,
  saveOutputDir: null,
  outputResolved: null,
  styleTheme: null,
  styleTelopType: null,
  styleTextPosition: null,
  styleFont: null,
  styleColor: null,
  styleStrokeColor: null,
  styleStrokeWidth: null,
  saveStyle: null,
  refreshPreview: null,
  previewMeta: null,
  previewImage: null,
  previewOverlay: null,
  previewOverlayText: null,
  status: null,
};

const api = async (url, options = {}) => {
  const res = await fetch(url, {
    headers: {'Content-Type': 'application/json'},
    ...options,
  });
  const data = await res.json().catch(() => ({}));
  if (!res.ok) {
    throw new Error(data.error || `HTTP ${res.status}`);
  }
  return data;
};

const showMessage = (text, type = 'info') => {
  if (!els.message) return;
  els.message.textContent = String(text || '');
  els.message.className = `message ${type}`;
};

const clearMessage = () => {
  if (!els.message) return;
  els.message.textContent = '';
  els.message.className = 'message hidden';
};

const setBusy = (button, busy) => {
  if (!button) return;
  if (!button.dataset.originalText) button.dataset.originalText = button.textContent;
  button.disabled = busy;
  button.textContent = busy ? '処理中...' : button.dataset.originalText;
};

const normalizeColorHex = (value, fallback = '#ffffff') => {
  const raw = String(value || '').trim();
  return /^#[0-9a-fA-F]{6}$/.test(raw) ? raw.toLowerCase() : fallback;
};

const normalizeStrokeWidth = (value, fallback = 0.8) => {
  const n = Number.parseFloat(String(value || ''));
  if (!Number.isFinite(n)) return fallback;
  const clamped = Math.max(0, Math.min(4, n));
  return Math.round(clamped * 10) / 10;
};

const setSelectOptions = (el, values, preferred = '') => {
  if (!el) return;
  const list = Array.from(new Set((values || []).map((v) => String(v || '').trim()).filter(Boolean)));
  el.innerHTML = '';
  list.forEach((v) => {
    const op = document.createElement('option');
    op.value = v;
    op.textContent = v;
    el.appendChild(op);
  });
  if (list.length === 0) return;
  el.value = list.includes(preferred) ? preferred : list[0];
};

const mapFontToCssFamily = (value) => {
  const raw = String(value || '');
  if (raw.includes('明朝')) return "'Yu Mincho', 'Hiragino Mincho ProN', serif";
  if (raw.includes('丸')) return "'M PLUS Rounded 1c', 'Hiragino Maru Gothic ProN', sans-serif";
  if (raw.includes('UD')) return "'BIZ UDPGothic', 'Yu Gothic UI', 'Meiryo', sans-serif";
  return "'Yu Gothic UI', 'Meiryo', sans-serif";
};

const mapThemeFontToOption = (value) => {
  const raw = String(value || '');
  if (raw.includes('明朝')) return '明朝';
  if (raw.includes('丸')) return '丸ゴシック';
  if (raw.includes('UD')) return 'UDゴシック';
  if (raw.includes('太')) return '太ゴシック';
  if (raw.includes('角')) return '角ゴシック';
  if (raw.includes('ゴシック')) return '標準';
  return '';
};

const THEME_PREVIEW_STYLES = {
  clean: {
    overlay: 'linear-gradient(180deg, rgba(0,0,0,0.36) 0%, rgba(0,0,0,0.12) 100%)',
    boxBg: 'linear-gradient(140deg, rgba(255,255,255,0.22) 0%, rgba(220,235,255,0.20) 100%)',
    boxBorder: '1px solid rgba(255,255,255,0.52)',
  },
  pop: {
    overlay: 'linear-gradient(180deg, rgba(56,25,102,0.30) 0%, rgba(182,54,140,0.16) 100%)',
    boxBg: 'linear-gradient(135deg, rgba(255,129,167,0.34) 0%, rgba(111,201,255,0.30) 100%)',
    boxBorder: '1px solid rgba(255,255,255,0.62)',
  },
  cinematic: {
    overlay: 'linear-gradient(180deg, rgba(0,0,0,0.62) 0%, rgba(0,0,0,0.26) 100%)',
    boxBg: 'linear-gradient(180deg, rgba(8,8,8,0.84) 0%, rgba(22,22,22,0.72) 100%)',
    boxBorder: '1px solid rgba(255,255,255,0.24)',
  },
  noir: {
    overlay: 'linear-gradient(180deg, rgba(0,0,0,0.68) 0%, rgba(0,0,0,0.30) 100%)',
    boxBg: 'linear-gradient(180deg, rgba(18,18,18,0.82) 0%, rgba(10,10,10,0.72) 100%)',
    boxBorder: '1px solid rgba(255,255,255,0.16)',
  },
  sunset: {
    overlay: 'linear-gradient(180deg, rgba(92,34,22,0.44) 0%, rgba(145,58,34,0.20) 100%)',
    boxBg: 'linear-gradient(135deg, rgba(255,171,124,0.36) 0%, rgba(255,118,84,0.24) 100%)',
    boxBorder: '1px solid rgba(255,233,214,0.58)',
  },
  aqua: {
    overlay: 'linear-gradient(180deg, rgba(0,42,56,0.42) 0%, rgba(0,86,108,0.18) 100%)',
    boxBg: 'linear-gradient(135deg, rgba(162,247,255,0.26) 0%, rgba(98,219,255,0.18) 100%)',
    boxBorder: '1px solid rgba(219,251,255,0.56)',
  },
};

const resolvePosClass = (pos) => {
  const raw = String(pos || '');
  if (raw === (state.textPositions[0] || '上')) return 'pos-top';
  if (raw === (state.textPositions[2] || '下')) return 'pos-bottom';
  return 'pos-center';
};

const renderOverlay = (scene) => {
  if (!ui.previewOverlay || !ui.previewOverlayText) return;
  if (!scene) {
    ui.previewOverlay.classList.add('hidden');
    return;
  }
  const text = String(scene.textRaw || '').trim() || PREVIEW_FALLBACK_TEXT;
  const color = normalizeColorHex(scene.color, '#ffffff');
  const strokeColor = normalizeColorHex(scene.strokeColor, '#121212');
  const strokeWidth = normalizeStrokeWidth(scene.strokeWidth, 0.8);
  const themeName = String(scene.theme || '').trim().toLowerCase();
  const themePreset = THEME_PREVIEW_STYLES[themeName] || THEME_PREVIEW_STYLES.clean;

  ui.previewOverlay.classList.remove('hidden', 'pos-top', 'pos-center', 'pos-bottom');
  ui.previewOverlay.classList.add(resolvePosClass(scene.textPosition));
  ui.previewOverlay.style.background = themePreset.overlay;

  const t = ui.previewOverlayText;
  t.textContent = text;
  t.style.color = color;
  t.style.fontFamily = mapFontToCssFamily(scene.font);
  t.style.webkitTextStroke = `${strokeWidth}px ${strokeColor}`;
  t.style.textShadow = '0 2px 6px rgba(0,0,0,0.35)';
  t.style.maxWidth = '100%';
  t.style.width = '100%';
  t.style.margin = '0';
  t.style.display = 'block';
  t.style.fontWeight = '';
  t.style.letterSpacing = '';

  const telopType = String(scene.telopType || '');
  if (telopType === '帯') {
    t.style.background = 'rgba(0,0,0,0.38)';
    t.style.border = '0';
    t.style.borderRadius = '10px';
    t.style.padding = '10px 12px';
    t.style.maxWidth = '100%';
  } else if (telopType.includes('吹き出し')) {
    t.style.background = 'rgba(255,255,255,0.18)';
    t.style.border = '1px solid rgba(255,255,255,0.42)';
    t.style.borderRadius = '14px';
    t.style.padding = '10px 12px';
    t.style.maxWidth = '100%';
  } else if (telopType.includes('角丸ラベル')) {
    t.style.background = 'rgba(24,24,24,0.52)';
    t.style.border = '1px solid rgba(255,255,255,0.36)';
    t.style.borderRadius = '22px';
    t.style.padding = '8px 14px';
    t.style.maxWidth = '90%';
    t.style.width = 'fit-content';
    t.style.margin = '0 auto';
    t.style.display = 'inline-block';
  } else if (telopType.includes('強調')) {
    t.style.background = 'linear-gradient(135deg, rgba(255,120,80,0.42), rgba(255,208,80,0.38))';
    t.style.border = '1px solid rgba(255,255,255,0.50)';
    t.style.borderRadius = '16px';
    t.style.padding = '10px 14px';
    t.style.maxWidth = '100%';
    t.style.fontWeight = '800';
    t.style.letterSpacing = '0.02em';
    t.style.webkitTextStroke = `${Math.max(strokeWidth, 1.2)}px ${strokeColor}`;
  } else if (telopType.includes('ネオン')) {
    t.style.background = 'rgba(12,18,32,0.34)';
    t.style.border = '1px solid rgba(136,246,255,0.78)';
    t.style.borderRadius = '14px';
    t.style.padding = '10px 14px';
    t.style.maxWidth = '100%';
    t.style.fontWeight = '700';
    t.style.letterSpacing = '0.02em';
    t.style.webkitTextStroke = `${Math.max(strokeWidth, 1)}px #7cf4ff`;
    t.style.textShadow = '0 0 10px rgba(124,244,255,0.88), 0 0 20px rgba(124,244,255,0.52)';
  } else if (telopType.includes('ガラス')) {
    t.style.background = 'linear-gradient(135deg, rgba(255,255,255,0.26), rgba(210,225,255,0.18))';
    t.style.border = '1px solid rgba(255,255,255,0.58)';
    t.style.borderRadius = '16px';
    t.style.padding = '10px 14px';
    t.style.maxWidth = '100%';
    t.style.fontWeight = '700';
    t.style.letterSpacing = '0.01em';
  } else if (telopType.includes('アウトライン')) {
    t.style.background = 'transparent';
    t.style.border = '0';
    t.style.borderRadius = '0';
    t.style.padding = '0';
    t.style.maxWidth = '100%';
    t.style.fontWeight = '800';
    t.style.letterSpacing = '0.02em';
    t.style.webkitTextStroke = `${Math.max(strokeWidth, 2)}px ${strokeColor}`;
    t.style.textShadow = '0 3px 12px rgba(0,0,0,0.45)';
  } else {
    t.style.background = themePreset.boxBg;
    t.style.border = themePreset.boxBorder;
    t.style.borderRadius = '16px';
    t.style.padding = '10px 14px';
    t.style.maxWidth = '100%';
    t.style.fontWeight = '';
    t.style.letterSpacing = '';
    t.style.webkitTextStroke = `${strokeWidth}px ${strokeColor}`;
  }
};

const pickPreviewStory = () => {
  if (state.previewStoryId) {
    const found = state.stories.find((s) => String(s.storyId || '') === state.previewStoryId);
    if (found) return found;
  }
  const candidates = [...state.stories].reverse();
  return candidates.find((s) => String(s.imageFileId || '').trim()) || candidates[0] || null;
};

const pickSeedScene = (scenes) => {
  const list = Array.isArray(scenes) ? scenes : [];
  return list.find((s) => String(s?.textRaw || '').trim()) || list[0] || null;
};

const loadSceneSeed = async (storyId) => {
  const key = String(storyId || '');
  if (!key) return null;
  if (state.sceneCache.has(key)) return state.sceneCache.get(key);
  try {
    const data = await api(`/api/stories/${encodeURIComponent(key)}/scenes`);
    const seed = pickSeedScene(data.scenes || []);
    if (seed) state.sceneCache.set(key, seed);
    return seed;
  } catch {
    return {
      sceneNo: 1,
      textRaw: PREVIEW_FALLBACK_TEXT,
      telopType: state.telopTypes[0] || '標準',
      textPosition: state.textPositions[1] || '中央',
      font: FONT_OPTIONS[0],
      color: '#ffffff',
      strokeColor: '#121212',
      strokeWidth: 0.8,
    };
  }
};

const buildScene = (baseScene) => ({
  sceneNo: Number(baseScene?.sceneNo || 1),
  textRaw: String(baseScene?.textRaw || ''),
  theme: ui.styleTheme?.value || baseScene?.theme || 'clean',
  telopType: ui.styleTelopType?.value || baseScene?.telopType || state.telopTypes[0] || '標準',
  textPosition: ui.styleTextPosition?.value || baseScene?.textPosition || state.textPositions[1] || '中央',
  font: ui.styleFont?.value || baseScene?.font || FONT_OPTIONS[0],
  color: normalizeColorHex(ui.styleColor?.value || baseScene?.color, '#ffffff'),
  strokeColor: normalizeColorHex(ui.styleStrokeColor?.value || baseScene?.strokeColor, '#121212'),
  strokeWidth: normalizeStrokeWidth(ui.styleStrokeWidth?.value ?? baseScene?.strokeWidth, 0.8),
});

const loadImageWithTimeout = (img, src, timeoutMs = 12000) =>
  new Promise((resolve, reject) => {
    let done = false;
    const finish = (err) => {
      if (done) return;
      done = true;
      clearTimeout(timer);
      img.onload = null;
      img.onerror = null;
      if (err) reject(err);
      else resolve();
    };
    const timer = window.setTimeout(() => finish(new Error('画像の読み込みがタイムアウトしました。')), timeoutMs);
    img.onload = () => finish();
    img.onerror = () => finish(new Error('画像の読み込みに失敗しました。'));
    img.src = src;
  });

const updateOverlayOnly = async () => {
  const story = pickPreviewStory();
  const storyId = String(story?.storyId || '').trim();
  if (!storyId) {
    renderOverlay({textRaw: PREVIEW_FALLBACK_TEXT, telopType: '標準', textPosition: '中央', font: FONT_OPTIONS[0]});
    if (ui.previewMeta) ui.previewMeta.textContent = '対象: -';
    return;
  }
  const baseScene = state.sceneCache.get(storyId) || {sceneNo: 1, textRaw: PREVIEW_FALLBACK_TEXT};
  const scene = buildScene(baseScene);
  renderOverlay(scene);
  if (ui.previewMeta) ui.previewMeta.textContent = `対象: ${storyId} / シーン${scene.sceneNo}`;
};

const refreshPreviewImage = async ({silent = false} = {}) => {
  if (!ui.previewImage || !ui.previewMeta) return;
  if (state.auth?.needsAuth) {
    ui.previewImage.classList.add('hidden');
    renderOverlay(null);
    ui.previewMeta.textContent = 'Googleログイン後にプレビューを表示できます。';
    return;
  }
  if (state.previewInFlight) {
    state.previewPending = true;
    return;
  }

  const story = pickPreviewStory();
  const storyId = String(story?.storyId || '').trim();
  const imageFileId = String(story?.imageFileId || '').trim();
  if (!storyId || !imageFileId) {
    ui.previewImage.classList.add('hidden');
    renderOverlay(null);
    ui.previewMeta.textContent = '対象ストーリーまたは画像IDがありません。';
    return;
  }

  state.previewInFlight = true;
  try {
    setBusy(ui.refreshPreview, true);
    state.previewStoryId = storyId;

    const seed = await loadSceneSeed(storyId);
    const scene = buildScene(seed);
    renderOverlay(scene);
    ui.previewMeta.textContent = `対象: ${storyId} / シーン${scene.sceneNo}`;

    const data = await api(`/api/stories/${encodeURIComponent(storyId)}/preview-image`, {
      method: 'POST',
      body: JSON.stringify({
        imageFileId,
        knownSourceKey: state.previewSourceKey,
      }),
    });

    if (!data.unchanged && data.imageSrc) {
      await loadImageWithTimeout(ui.previewImage, data.imageSrc);
      ui.previewImage.classList.remove('hidden');
    } else if (ui.previewImage.src) {
      ui.previewImage.classList.remove('hidden');
    }
    state.previewSourceKey = String(data.sourceKey || '');

    if (!silent) showMessage('右側画像を更新しました。', 'success');
  } catch (error) {
    const message = String(error?.message || error || '');
    ui.previewMeta.textContent = `プレビュー更新失敗: ${message}`;
    if (!silent) showMessage(message, 'error');
  } finally {
    state.previewInFlight = false;
    setBusy(ui.refreshPreview, false);
    if (state.previewPending) {
      state.previewPending = false;
      refreshPreviewImage({silent: true}).catch(() => {});
    }
  }
};

const applyThemePreset = () => {
  const current = String(ui.styleTheme?.value || '');
  if (!current) return;
  const preset = (state.themes || []).find((t) => String(t?.themeName || '') === current);
  if (!preset) return;
  if (ui.styleColor && preset.textColor) ui.styleColor.value = normalizeColorHex(preset.textColor, '#ffffff');
  if (ui.styleStrokeColor && preset.strokeColor) ui.styleStrokeColor.value = normalizeColorHex(preset.strokeColor, '#121212');
  const font = mapThemeFontToOption(preset.fontFamily);
  if (font && ui.styleFont) ui.styleFont.value = font;
};

const renderJobs = () => {
  if (!els.jobsList || !els.jobsEmpty) return;
  els.jobsList.innerHTML = '';
  const jobs = Array.isArray(state.jobs) ? [...state.jobs] : [];
  if (jobs.length === 0) {
    els.jobsEmpty.classList.remove('hidden');
    return;
  }
  els.jobsEmpty.classList.add('hidden');
  jobs
    .sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')))
    .slice(0, 40)
    .forEach((job, idx) => {
      const card = document.createElement('div');
      card.className = 'job-card';
      card.innerHTML =
        `<div class="job-head"><div class="job-head-left"><span class="job-order">${idx + 1}</span><strong>${job.storyId || '-'}</strong></div>` +
        `<span class="status-chip">${job.status || '-'}</span></div>` +
        `<div class="muted">更新: ${job.updatedAt || '-'} / 作成: ${job.createdAt || '-'}</div>`;
      els.jobsList.appendChild(card);
    });
};

const renderSheetLocation = () => {
  if (!els.sheetLocation) return;
  const id = String(state.config?.spreadsheetId || '').trim();
  if (!id) {
    els.sheetLocation.textContent = 'スプレッドシート情報を読み込めませんでした。';
    return;
  }
  const url = `https://docs.google.com/spreadsheets/d/${id}/edit`;
  const outputLocalDir = String(state.config?.outputLocalDir || 'output').trim() || 'output';
  const outputResolvedDir = String(state.config?.outputResolvedDir || '').trim();
  if (els.spreadsheetIdInput) els.spreadsheetIdInput.value = id;
  els.sheetLocation.innerHTML =
    `<div>シートID: <code>${id}</code></div>` +
    `<div>シートURL: <a href="${url}" target="_blank" rel="noopener">${url}</a></div>` +
    `<div>ローカル出力先: <code>${outputLocalDir}</code>${outputResolvedDir ? ` (${outputResolvedDir})` : ''}</div>`;
};

const renderSimpleStatus = () => {
  if (!ui.status) return;
  const waiting = state.jobs.filter((j) => String(j.status || '') === '待機中').length;
  const rendering = state.jobs.filter((j) => String(j.status || '') === '書き出し中').length;
  const done = state.jobs.filter((j) => String(j.status || '') === '完了').length;
  ui.status.textContent = `ストーリー ${state.stories.length}件 / 待機 ${waiting}件 / 処理中 ${rendering}件 / 完了 ${done}件`;
};

const openSpreadsheet = () => {
  const id = String(state.config?.spreadsheetId || '').trim();
  if (!id) {
    showMessage('スプレッドシートIDが未設定です。', 'error');
    return;
  }
  window.open(`https://docs.google.com/spreadsheets/d/${id}/edit`, '_blank', 'noopener');
};

const saveStyle = async () => {
  const renderStyle = {
    theme: ui.styleTheme?.value || 'clean',
    telopType: ui.styleTelopType?.value || state.telopTypes[0] || '標準',
    textPosition: ui.styleTextPosition?.value || state.textPositions[1] || '中央',
    font: ui.styleFont?.value || FONT_OPTIONS[0],
    color: normalizeColorHex(ui.styleColor?.value, '#ffffff'),
    strokeColor: normalizeColorHex(ui.styleStrokeColor?.value, '#121212'),
    strokeWidth: normalizeStrokeWidth(ui.styleStrokeWidth?.value, 0.8),
  };
  setBusy(ui.saveStyle, true);
  try {
    const res = await api('/api/config', {method: 'PUT', body: JSON.stringify({renderStyle})});
    state.config = {...(state.config || {}), ...(res.config || {}), renderStyle};
    showMessage('スタイルを保存しました。', 'success');
  } catch (error) {
    showMessage(error.message, 'error');
  } finally {
    setBusy(ui.saveStyle, false);
  }
};

const saveOutputDir = async () => {
  const outputLocalDir = String(ui.outputDir?.value || '').trim() || 'output';
  setBusy(ui.saveOutputDir, true);
  try {
    const res = await api('/api/config', {method: 'PUT', body: JSON.stringify({outputLocalDir})});
    state.config = {...(state.config || {}), ...(res.config || {})};
    renderSheetLocation();
    renderOutputDirControls();
    showMessage('ローカル保存先を保存しました。', 'success');
  } catch (error) {
    showMessage(error.message, 'error');
  } finally {
    setBusy(ui.saveOutputDir, false);
  }
};

const browseOutputDir = async () => {
  setBusy(ui.browseOutputDir, true);
  try {
    showMessage('エクスプローラーでフォルダ選択を開いています。', 'info');
    const currentDir = String(ui.outputDir?.value || state.config?.outputLocalDir || 'output').trim();
    const res = await api('/api/dialog/select-output-dir', {
      method: 'POST',
      body: JSON.stringify({currentDir}),
    });
    if (res.canceled) {
      showMessage('フォルダ選択をキャンセルしました。', 'info');
      return;
    }
    const selected = String(res.outputLocalDir || '').trim();
    if (!selected) {
      throw new Error('フォルダを取得できませんでした。');
    }
    if (ui.outputDir) ui.outputDir.value = selected;
    await saveOutputDir();
  } catch (error) {
    showMessage(error.message, 'error');
  } finally {
    setBusy(ui.browseOutputDir, false);
  }
};

const startGoogleAuth = async () => {
  setBusy(ui.googleAuth, true);
  try {
    const status = await api('/api/auth/google/status');
    if (!status.ready) {
      throw new Error('Googleログイン設定が見つかりません。');
    }
    const data = await api('/api/auth/google/start', {method: 'POST'});
    const popup = window.open(data.authUrl, 'google-auth', 'width=560,height=760');
    if (!popup) throw new Error('ポップアップがブロックされました。');
    showMessage('Googleログイン画面を開きました。認証後に再読込します。', 'info');
  } catch (error) {
    showMessage(error.message, 'error');
  } finally {
    setBusy(ui.googleAuth, false);
  }
};

const mountSimplePanel = () => {
  if (ui.panel) return;
  const header = document.querySelector('.header');
  if (!header) return;

  const panel = document.createElement('section');
  panel.className = 'panel';
  panel.id = 'simple-main-panel';
  panel.innerHTML =
    '<h2>メイン操作</h2>' +
    '<div class="simple-main-layout">' +
  '  <div class="simple-main-left">' +
    '    <div class="button-row">' +
    '      <button id="simple-start-all">シート内容で一括処理開始</button>' +
    '      <button id="simple-google-auth" class="secondary">Googleログイン</button>' +
    '      <button id="simple-open-sheet" class="secondary">スプレッドシートを開く</button>' +
    '      <button id="simple-reload" class="secondary">再読み込み</button>' +
    '    </div>' +
    '    <div class="range-box">' +
    '      <div class="range-head">処理範囲(ストーリーID)</div>' +
    '      <div class="range-grid">' +
    '        <div class="form-row"><label>開始ID</label><input id="simple-range-start-id" placeholder="TTK_20260210_001" /></div>' +
    '        <div class="form-row"><label>終了ID</label><input id="simple-range-end-id" placeholder="TTK_20260210_050" /></div>' +
    '      </div>' +
    '      <div class="button-row"><button id="simple-start-range" class="secondary">範囲指定で処理開始</button></div>' +
    '      <div class="field-hint">開始IDまたは終了IDだけでも指定可能です（片方未入力なら先頭/末尾まで）。</div>' +
    '    </div>' +
    '    <div class="output-box">' +
    '      <div class="output-head">ローカル保存先</div>' +
    '      <div class="form-row"><label>出力フォルダ（相対 or 絶対パス）</label><div class="input-with-button"><input id="simple-output-dir" placeholder="output または C:\\\\TikTok\\\\output" /><button id="simple-browse-output-dir" type="button" class="secondary">参照</button></div></div>' +
    '      <div class="button-row"><button id="simple-save-output-dir" class="secondary">保存先を保存</button></div>' +
    '      <div id="simple-output-resolved" class="field-hint"></div>' +
    '    </div>' +
    '    <div class="scene-style-box">' +
      '      <div class="scene-style-head">共通スタイル(全件適用)</div>' +
    '      <div class="grid-3">' +
    '        <div class="form-row"><label>テーマ</label><select id="simple-style-theme"></select></div>' +
    '        <div class="form-row"><label>テロップ種類</label><select id="simple-style-telop-type"></select></div>' +
    '        <div class="form-row"><label>テキスト位置</label><select id="simple-style-text-position"></select></div>' +
    '      </div>' +
    '      <div class="grid-3">' +
    '        <div class="form-row"><label>フォント</label><select id="simple-style-font"></select></div>' +
    '        <div class="form-row"><label>文字色</label><input id="simple-style-color" type="color" value="#ffffff" /></div>' +
    '        <div class="form-row"><label>フチ色</label><input id="simple-style-stroke-color" type="color" value="#121212" /></div>' +
    '      </div>' +
    '      <div class="grid-3">' +
    '        <div class="form-row"><label>フチ太さ</label><select id="simple-style-stroke-width"></select></div>' +
    '      </div>' +
    '      <div class="button-row"><button id="simple-save-style" class="secondary">スタイル保存</button></div>' +
    '    </div>' +
    '    <div class="field-hint">スタイルは即時反映されます。画像を再取得したい時だけ右側の更新を押してください。</div>' +
    '    <div id="simple-main-status" class="field-hint"></div>' +
    '  </div>' +
    '  <div class="simple-main-right">' +
    '    <div class="preview-wrap simple-preview-wrap">' +
    '      <div class="preview-head"><strong>プレビュー(1枚)</strong><span id="simple-preview-meta" class="muted">対象: -</span></div>' +
    '      <div class="button-row"><button id="simple-refresh-preview" class="secondary">右側画像更新</button></div>' +
    '      <div id="simple-preview-stage" class="simple-preview-stage">' +
    '        <img id="simple-preview-image" class="preview-image hidden" alt="プレビュー" />' +
    '        <div id="simple-preview-overlay" class="simple-preview-overlay hidden pos-center">' +
    '          <div id="simple-preview-overlay-text" class="simple-preview-overlay-text"></div>' +
    '        </div>' +
    '      </div>' +
    '    </div>' +
    '  </div>' +
    '</div>';

  header.insertAdjacentElement('afterend', panel);
  ui.panel = panel;
  ui.startAll = panel.querySelector('#simple-start-all');
  ui.startRange = panel.querySelector('#simple-start-range');
  ui.googleAuth = panel.querySelector('#simple-google-auth');
  ui.openSheet = panel.querySelector('#simple-open-sheet');
  ui.reload = panel.querySelector('#simple-reload');
  ui.rangeStartId = panel.querySelector('#simple-range-start-id');
  ui.rangeEndId = panel.querySelector('#simple-range-end-id');
  ui.outputDir = panel.querySelector('#simple-output-dir');
  ui.browseOutputDir = panel.querySelector('#simple-browse-output-dir');
  ui.saveOutputDir = panel.querySelector('#simple-save-output-dir');
  ui.outputResolved = panel.querySelector('#simple-output-resolved');
  ui.styleTheme = panel.querySelector('#simple-style-theme');
  ui.styleTelopType = panel.querySelector('#simple-style-telop-type');
  ui.styleTextPosition = panel.querySelector('#simple-style-text-position');
  ui.styleFont = panel.querySelector('#simple-style-font');
  ui.styleColor = panel.querySelector('#simple-style-color');
  ui.styleStrokeColor = panel.querySelector('#simple-style-stroke-color');
  ui.styleStrokeWidth = panel.querySelector('#simple-style-stroke-width');
  ui.saveStyle = panel.querySelector('#simple-save-style');
  ui.refreshPreview = panel.querySelector('#simple-refresh-preview');
  ui.previewMeta = panel.querySelector('#simple-preview-meta');
  ui.previewImage = panel.querySelector('#simple-preview-image');
  ui.previewOverlay = panel.querySelector('#simple-preview-overlay');
  ui.previewOverlayText = panel.querySelector('#simple-preview-overlay-text');
  ui.status = panel.querySelector('#simple-main-status');

  const autoControls = [ui.styleTheme, ui.styleTelopType, ui.styleTextPosition, ui.styleFont, ui.styleColor, ui.styleStrokeColor, ui.styleStrokeWidth].filter(Boolean);
  autoControls.forEach((el) => {
    const ev = String(el.type || '').toLowerCase() === 'color' ? 'input' : 'change';
    el.addEventListener(ev, () => {
      if (el === ui.styleTheme) applyThemePreset();
      updateOverlayOnly().catch(() => {});
    });
  });

  ui.saveStyle?.addEventListener('click', () => saveStyle());
  ui.browseOutputDir?.addEventListener('click', () => browseOutputDir());
  ui.saveOutputDir?.addEventListener('click', () => saveOutputDir());
  ui.refreshPreview?.addEventListener('click', () => refreshPreviewImage());
  ui.reload?.addEventListener('click', () => bootstrap());
  ui.openSheet?.addEventListener('click', () => openSpreadsheet());
  ui.googleAuth?.addEventListener('click', () => startGoogleAuth());
  ui.startAll?.addEventListener('click', async () => {
    setBusy(ui.startAll, true);
    try {
      const r = await api('/api/batch/enqueue-all-from-sheet', {method: 'POST'});
      showMessage(`キュー投入: ${r.queuedCount || 0}件`, 'success');
      await bootstrap();
    } catch (error) {
      showMessage(error.message, 'error');
    } finally {
      setBusy(ui.startAll, false);
    }
  });
  ui.startRange?.addEventListener('click', async () => {
    const startStoryId = String(ui.rangeStartId?.value || '').trim();
    const endStoryId = String(ui.rangeEndId?.value || '').trim();
    if (!startStoryId && !endStoryId) {
      showMessage('開始IDか終了IDを入力してください。', 'error');
      return;
    }
    setBusy(ui.startRange, true);
    try {
      const r = await api('/api/batch/enqueue-range-from-sheet', {
        method: 'POST',
        body: JSON.stringify({startStoryId, endStoryId}),
      });
      const startLabel = r.rangeStartId || startStoryId || '先頭';
      const endLabel = r.rangeEndId || endStoryId || '末尾';
      if (Number(r.queuedCount || 0) > 0) {
        showMessage(`範囲キュー投入: ${r.queuedCount || 0}件 (${startLabel} 〜 ${endLabel})`, 'success');
      } else {
        const skippedHint = Array.isArray(r.skipped) && r.skipped.length > 0 ? ` / スキップ理由: ${String(r.skipped[0].reason || '-')}` : '';
        showMessage(`範囲対象 ${r.considered || 0}件、投入0件でした (${startLabel} 〜 ${endLabel})${skippedHint}`, 'info');
      }
      await bootstrap();
    } catch (error) {
      showMessage(error.message, 'error');
    } finally {
      setBusy(ui.startRange, false);
    }
  });
};

const applyAuthState = () => {
  const disabled = Boolean(state.auth?.needsAuth);
  [
    ui.startAll,
    ui.startRange,
    ui.rangeStartId,
    ui.rangeEndId,
    ui.styleTheme,
    ui.styleTelopType,
    ui.styleTextPosition,
    ui.styleFont,
    ui.styleColor,
    ui.styleStrokeColor,
    ui.styleStrokeWidth,
    ui.saveStyle,
    ui.refreshPreview,
  ].forEach((el) => {
    if (el) el.disabled = disabled;
  });
  if (ui.googleAuth) ui.googleAuth.disabled = false;
};

const renderStyleControls = () => {
  const style = state.config?.renderStyle || {};
  const themeNames = state.themes.length > 0 ? state.themes.map((t) => t.themeName) : ['clean'];
  setSelectOptions(ui.styleTheme, themeNames, style.theme || 'clean');
  setSelectOptions(ui.styleTelopType, state.telopTypes, style.telopType || state.telopTypes[0]);
  setSelectOptions(ui.styleTextPosition, state.textPositions, style.textPosition || state.textPositions[1]);
  setSelectOptions(ui.styleFont, FONT_OPTIONS, style.font || FONT_OPTIONS[0]);
  setSelectOptions(ui.styleStrokeWidth, STROKE_WIDTH_OPTIONS, String(normalizeStrokeWidth(style.strokeWidth, 0.8)));
  if (ui.styleColor) ui.styleColor.value = normalizeColorHex(style.color, '#ffffff');
  if (ui.styleStrokeColor) ui.styleStrokeColor.value = normalizeColorHex(style.strokeColor, '#121212');
};

const renderOutputDirControls = () => {
  const outputLocalDir = String(state.config?.outputLocalDir || 'output').trim() || 'output';
  const outputResolvedDir = String(state.config?.outputResolvedDir || '').trim();
  if (ui.outputDir) ui.outputDir.value = outputLocalDir;
  if (ui.outputResolved) {
    ui.outputResolved.textContent = outputResolvedDir ? `現在の保存先: ${outputResolvedDir}` : '';
  }
};

const hideAdvancedPanels = () => {
  const panels = Array.from(document.querySelectorAll('.panel'));
  panels.forEach((panel) => {
    if (panel.id === 'simple-main-panel') return;
    const title = panel.querySelector('h2')?.textContent?.trim();
    if (title === 'スプレッドシートリンク' || title === 'ジョブ進捗') return;
    panel.classList.add('hidden');
  });
};

const bootstrap = async () => {
  clearMessage();
  try {
    const data = await api('/api/bootstrap');
    state.config = data.config || {renderStyle: {}};
    state.stories = data.stories || [];
    state.jobs = data.jobs || [];
    state.themes = data.themes || [];
    state.telopTypes = data.telopTypes || state.telopTypes;
    state.textPositions = data.textPositions || state.textPositions;
    state.auth = {
      ...state.auth,
      ...(data.auth || {}),
      needsAuth: Boolean(data.needsAuth),
    };

    renderSheetLocation();
    renderOutputDirControls();
    renderStyleControls();
    renderJobs();
    renderSimpleStatus();
    applyAuthState();
    hideAdvancedPanels();
    await updateOverlayOnly();
    await refreshPreviewImage({silent: true});

    if (state.auth.needsAuth && state.auth.message) {
      showMessage(state.auth.message, 'info');
    }
  } catch (error) {
    showMessage(error.message, 'error');
  }
};

const init = async () => {
  mountSimplePanel();
  els.openSheetTop?.addEventListener('click', () => openSpreadsheet());
  window.addEventListener('message', (event) => {
    if (!event?.data || event.data.type !== 'google-auth-complete') return;
    if (event.data.ok) {
      showMessage('Googleログイン完了。再読み込みします。', 'success');
      bootstrap().catch(() => {});
    }
  });
  await bootstrap();
};

init().catch((error) => {
  showMessage(error.message, 'error');
});
