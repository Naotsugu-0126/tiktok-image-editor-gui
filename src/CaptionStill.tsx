import React from 'react';
import {AbsoluteFill, Img} from 'remotion';
import {loadFont} from '@remotion/google-fonts/YuseiMagic';

type CaptionStillProps = {
  imageSrc: string;
  text: string;
  theme?: string;
  font?: string;
  color?: string;
  strokeColor?: string;
  strokeWidth?: number | string;
  telopType?: string;
  textPosition?: string;
  previewMode?: 'normal' | 'fast';
};

type ThemeName = 'clean' | 'pop' | 'cinematic' | 'noir' | 'sunset' | 'aqua';
type TelopType = '標準' | '帯' | '吹き出し' | '角丸ラベル' | '強調' | 'ネオン' | 'ガラス' | 'アウトライン';
type TextPosition = '上' | '中央' | '下';

const {fontFamily} = loadFont();

const chooseBreakIndex = (text: string): number | null => {
  const minChunk = 6;
  if (text.length < minChunk * 2) return null;

  const midpoint = Math.floor(text.length / 2);
  const tailPunctuation = new Set(['。', '、', '！', '？']);
  const particles = new Set(['は', 'が', 'を', 'に', 'と', 'で', 'へ', 'の', 'や', 'も']);
  let bestIndex = -1;
  let bestScore = -Infinity;

  for (let i = minChunk; i <= text.length - minChunk; i += 1) {
    const prev = text[i - 1];
    const next = text[i];
    let score = -Math.abs(i - midpoint) * 2;

    if (tailPunctuation.has(prev)) score += 6;
    if (particles.has(prev)) score += 2;
    if (particles.has(next)) score -= 2;
    if (next === 'っ' || next === 'ゃ' || next === 'ゅ' || next === 'ょ') score -= 3;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex > 0 ? bestIndex : null;
};

const formatCaption = (rawText: string): string => {
  const text = rawText.replace(/\\n/g, '\n').trim();
  if (text.includes('\n') || text.length <= 14) return text;
  const breakIndex = chooseBreakIndex(text);
  if (!breakIndex) return text;
  return `${text.slice(0, breakIndex)}\n${text.slice(breakIndex)}`;
};

const resolveTheme = (value: string | undefined): ThemeName => {
  const normalized = String(value ?? 'clean').trim().toLowerCase();
  if (normalized === 'pop') return 'pop';
  if (normalized === 'cinematic') return 'cinematic';
  if (normalized === 'noir') return 'noir';
  if (normalized === 'sunset') return 'sunset';
  if (normalized === 'aqua') return 'aqua';
  return 'clean';
};

const resolveTelopType = (value: string | undefined): TelopType => {
  const normalized = String(value ?? '標準').trim();
  if (normalized === '帯') return '帯';
  if (normalized === '吹き出し') return '吹き出し';
  if (normalized === '角丸ラベル') return '角丸ラベル';
  if (normalized === '強調') return '強調';
  if (normalized === 'ネオン') return 'ネオン';
  if (normalized === 'ガラス') return 'ガラス';
  if (normalized === 'アウトライン') return 'アウトライン';
  return '標準';
};

const resolveTextPosition = (value: string | undefined): TextPosition => {
  const normalized = String(value ?? '中央').trim();
  if (normalized === '上') return '上';
  if (normalized === '下') return '下';
  return '中央';
};

const resolveFontFamily = (font: string | undefined) => {
  const normalized = String(font ?? '').trim();
  if (normalized === '丸ゴシック' || normalized === '手書き') {
    return `'${fontFamily}', 'Hiragino Maru Gothic ProN', 'Yu Gothic UI', 'Meiryo', sans-serif`;
  }
  if (normalized === '明朝') {
    return `'Hiragino Mincho ProN', 'Yu Mincho', 'MS PMincho', serif`;
  }
  if (normalized === '太ゴシック') {
    return `'BIZ UDPGothic', 'Hiragino Kaku Gothic StdN', 'Yu Gothic', 'Meiryo', sans-serif`;
  }
  if (normalized === '角ゴシック') {
    return `'Hiragino Kaku Gothic ProN', 'Yu Gothic', 'Meiryo', sans-serif`;
  }
  if (normalized === 'UDゴシック') {
    return `'BIZ UDPGothic', 'Yu Gothic UI', 'Meiryo', sans-serif`;
  }
  return `'${fontFamily}', 'Hiragino Kaku Gothic ProN', 'Yu Gothic UI', 'Meiryo', sans-serif`;
};

const resolveFontWeight = (font: string | undefined): React.CSSProperties['fontWeight'] => {
  const normalized = String(font ?? '').trim();
  if (normalized === '太ゴシック') return 700;
  if (normalized === '明朝') return 600;
  return 500;
};

const resolveStrokeColor = (value: string | undefined, fallback = 'rgba(18, 18, 18, 0.16)') => {
  const normalized = String(value ?? '').trim();
  if (/^#[0-9a-fA-F]{6}$/.test(normalized)) return normalized;
  return fallback;
};

const resolveStrokeWidth = (value: number | string | undefined, fallback = 0.8) => {
  const parsed = Number.parseFloat(String(value ?? ''));
  if (!Number.isFinite(parsed)) return fallback;
  const clamped = Math.max(0, Math.min(4, parsed));
  return Math.round(clamped * 10) / 10;
};

const themeStyleMap: Record<ThemeName, {
  overlay: string;
  fastOverlay: string;
  boxBg: string;
  border: string;
  shadow: string;
  defaultColor: string;
}> = {
  clean: {
    overlay: 'linear-gradient(180deg, rgba(0,0,0,0.46) 0%, rgba(0,0,0,0.18) 42%, rgba(0,0,0,0.06) 100%)',
    fastOverlay: 'linear-gradient(180deg, rgba(0,0,0,0.36) 0%, rgba(0,0,0,0.12) 100%)',
    boxBg: 'linear-gradient(140deg, rgba(255,255,255,0.22) 0%, rgba(220,235,255,0.20) 100%)',
    border: '1px solid rgba(255,255,255,0.52)',
    shadow: '0 12px 34px rgba(0,0,0,0.20)',
    defaultColor: '#ffffff',
  },
  pop: {
    overlay: 'linear-gradient(180deg, rgba(37,18,72,0.34) 0%, rgba(208,50,101,0.20) 48%, rgba(0,0,0,0.04) 100%)',
    fastOverlay: 'linear-gradient(180deg, rgba(56,25,102,0.30) 0%, rgba(182,54,140,0.16) 100%)',
    boxBg: 'linear-gradient(135deg, rgba(255,129,167,0.34) 0%, rgba(111,201,255,0.30) 100%)',
    border: '1px solid rgba(255,255,255,0.62)',
    shadow: '0 14px 30px rgba(67,16,99,0.32)',
    defaultColor: '#ffffff',
  },
  cinematic: {
    overlay: 'linear-gradient(180deg, rgba(0,0,0,0.78) 0%, rgba(0,0,0,0.36) 42%, rgba(0,0,0,0.22) 100%)',
    fastOverlay: 'linear-gradient(180deg, rgba(0,0,0,0.62) 0%, rgba(0,0,0,0.26) 100%)',
    boxBg: 'linear-gradient(180deg, rgba(8,8,8,0.84) 0%, rgba(22,22,22,0.72) 100%)',
    border: '1px solid rgba(255,255,255,0.24)',
    shadow: '0 16px 42px rgba(0,0,0,0.62)',
    defaultColor: '#f9f8f4',
  },
  noir: {
    overlay: 'linear-gradient(180deg, rgba(2,2,2,0.86) 0%, rgba(8,8,8,0.42) 46%, rgba(0,0,0,0.24) 100%)',
    fastOverlay: 'linear-gradient(180deg, rgba(0,0,0,0.68) 0%, rgba(0,0,0,0.30) 100%)',
    boxBg: 'linear-gradient(180deg, rgba(18,18,18,0.82) 0%, rgba(10,10,10,0.72) 100%)',
    border: '1px solid rgba(255,255,255,0.16)',
    shadow: '0 16px 40px rgba(0,0,0,0.68)',
    defaultColor: '#f2f2f2',
  },
  sunset: {
    overlay: 'linear-gradient(180deg, rgba(73,26,18,0.56) 0%, rgba(165,74,44,0.28) 44%, rgba(0,0,0,0.16) 100%)',
    fastOverlay: 'linear-gradient(180deg, rgba(92,34,22,0.44) 0%, rgba(145,58,34,0.20) 100%)',
    boxBg: 'linear-gradient(135deg, rgba(255,171,124,0.36) 0%, rgba(255,118,84,0.24) 100%)',
    border: '1px solid rgba(255,233,214,0.58)',
    shadow: '0 14px 34px rgba(76,28,20,0.44)',
    defaultColor: '#fff7ee',
  },
  aqua: {
    overlay: 'linear-gradient(180deg, rgba(0,35,44,0.54) 0%, rgba(0,78,92,0.24) 44%, rgba(0,0,0,0.10) 100%)',
    fastOverlay: 'linear-gradient(180deg, rgba(0,42,56,0.42) 0%, rgba(0,86,108,0.18) 100%)',
    boxBg: 'linear-gradient(135deg, rgba(162,247,255,0.26) 0%, rgba(98,219,255,0.18) 100%)',
    border: '1px solid rgba(219,251,255,0.56)',
    shadow: '0 12px 30px rgba(0,42,56,0.34)',
    defaultColor: '#f4ffff',
  },
};

const positionStyleMap: Record<TextPosition, React.CSSProperties> = {
  上: {
    justifyContent: 'flex-start',
    padding: '116px 56px 0 56px',
  },
  中央: {
    justifyContent: 'center',
    padding: '56px',
  },
  下: {
    justifyContent: 'flex-end',
    padding: '0 56px 150px 56px',
  },
};

export const CaptionStill: React.FC<CaptionStillProps> = ({
  imageSrc,
  text,
  theme,
  font,
  color,
  strokeColor,
  strokeWidth,
  telopType,
  textPosition,
  previewMode,
}) => {
  const caption = formatCaption(text);
  const resolvedTheme = resolveTheme(theme);
  const resolvedTelopType = resolveTelopType(telopType);
  const resolvedPosition = resolveTextPosition(textPosition);
  const isFastPreview = previewMode === 'fast';

  const palette = themeStyleMap[resolvedTheme];
  const textColor = color && color.trim() ? color : palette.defaultColor;
  const textStrokeColor = resolveStrokeColor(strokeColor, 'rgba(18, 18, 18, 0.16)');
  const textStrokeWidth = resolveStrokeWidth(strokeWidth, 0.8);

  const charCount = caption.replace(/\s/g, '').length;
  const fontSize = charCount > 50 ? 50 : charCount > 35 ? 56 : 62;

  const boxStyle: React.CSSProperties = {
    width: '100%',
    maxWidth: 940,
    borderRadius: 34,
    background: palette.boxBg,
    border: palette.border,
    boxShadow: isFastPreview ? 'none' : palette.shadow,
    color: textColor,
    fontFamily: resolveFontFamily(font),
    fontSize,
    fontWeight: resolveFontWeight(font),
    lineHeight: 1.32,
    letterSpacing: '0.03em',
    textAlign: 'center',
    whiteSpace: 'pre-wrap',
    textShadow: isFastPreview ? 'none' : '0 4px 14px rgba(0, 0, 0, 0.46)',
    WebkitTextStroke: `${textStrokeWidth}px ${textStrokeColor}`,
  };

  if (resolvedTelopType === '帯') {
    boxStyle.maxWidth = 1080;
    boxStyle.borderRadius = 14;
    boxStyle.background = 'rgba(0, 0, 0, 0.58)';
    boxStyle.border = '1px solid rgba(255,255,255,0.22)';
  }

  if (resolvedTelopType === '吹き出し') {
    boxStyle.borderRadius = 42;
    boxStyle.background = 'rgba(255, 255, 255, 0.92)';
    boxStyle.border = '1px solid rgba(0,0,0,0.12)';
    boxStyle.color = color && color.trim() ? color : '#1e1e1e';
    boxStyle.textShadow = 'none';
    boxStyle.WebkitTextStroke = '0px transparent';
  }

  if (resolvedTelopType === '角丸ラベル') {
    boxStyle.maxWidth = 760;
    boxStyle.borderRadius = 999;
    boxStyle.background = 'rgba(18,18,18,0.56)';
    boxStyle.border = '1px solid rgba(255,255,255,0.36)';
    boxStyle.padding = '20px 28px 22px 28px';
  }

  if (resolvedTelopType === '強調') {
    boxStyle.borderRadius = 16;
    boxStyle.background = 'linear-gradient(135deg, rgba(255,120,80,0.42), rgba(255,208,80,0.38))';
    boxStyle.border = '1px solid rgba(255,255,255,0.5)';
    boxStyle.fontWeight = 800;
    boxStyle.letterSpacing = '0.02em';
    boxStyle.WebkitTextStroke = `${Math.max(textStrokeWidth, 1.2)}px ${textStrokeColor}`;
  }

  if (resolvedTelopType === 'ネオン') {
    boxStyle.borderRadius = 14;
    boxStyle.background = 'rgba(12,18,32,0.34)';
    boxStyle.border = '1px solid rgba(136,246,255,0.78)';
    boxStyle.fontWeight = 700;
    boxStyle.letterSpacing = '0.02em';
    boxStyle.WebkitTextStroke = `${Math.max(textStrokeWidth, 1)}px #7cf4ff`;
    boxStyle.textShadow = '0 0 10px rgba(124,244,255,0.88), 0 0 20px rgba(124,244,255,0.52)';
  }

  if (resolvedTelopType === 'ガラス') {
    boxStyle.borderRadius = 16;
    boxStyle.background = 'linear-gradient(135deg, rgba(255,255,255,0.26), rgba(210,225,255,0.18))';
    boxStyle.border = '1px solid rgba(255,255,255,0.58)';
    boxStyle.fontWeight = 700;
    boxStyle.letterSpacing = '0.01em';
  }

  if (resolvedTelopType === 'アウトライン') {
    boxStyle.background = 'transparent';
    boxStyle.border = '0';
    boxStyle.borderRadius = 0;
    boxStyle.fontWeight = 800;
    boxStyle.letterSpacing = '0.02em';
    boxStyle.WebkitTextStroke = `${Math.max(textStrokeWidth, 2)}px ${textStrokeColor}`;
    boxStyle.textShadow = '0 3px 12px rgba(0,0,0,0.45)';
  }

  return (
    <AbsoluteFill style={{backgroundColor: '#000'}}>
      <Img
        src={imageSrc}
        style={{
          width: '100%',
          height: '100%',
          objectFit: 'cover',
        }}
      />

      <AbsoluteFill
        style={{
          ...positionStyleMap[resolvedPosition],
          alignItems: 'center',
          background: isFastPreview ? palette.fastOverlay : palette.overlay,
        }}
      >
        <div
          style={{
            padding: '30px 36px 32px 36px',
            ...boxStyle,
          }}
        >
          {caption}
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
};
