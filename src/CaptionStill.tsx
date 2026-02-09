import React from 'react';
import {AbsoluteFill, Img} from 'remotion';
import {loadFont} from '@remotion/google-fonts/YuseiMagic';

type CaptionStillProps = {
  imageSrc: string;
  text: string;
};

const {fontFamily} = loadFont();

const chooseBreakIndex = (text: string): number | null => {
  const minChunk = 6;
  if (text.length < minChunk * 2) {
    return null;
  }

  const midpoint = Math.floor(text.length / 2);
  const tailPunctuation = new Set(['、', '。', '！', '？']);
  const commonParticles = new Set(['は', 'が', 'を', 'に', 'と', 'で', 'へ', 'の', 'も', 'や']);
  let bestIndex = -1;
  let bestScore = -Infinity;

  for (let i = minChunk; i <= text.length - minChunk; i += 1) {
    const prev = text[i - 1];
    const next = text[i];
    let score = -Math.abs(i - midpoint) * 2;

    if (tailPunctuation.has(prev)) score += 6;
    if (commonParticles.has(prev)) score += 2;
    if (commonParticles.has(next)) score -= 2;
    if (next === 'ゃ' || next === 'ゅ' || next === 'ょ' || next === 'っ') score -= 3;

    if (score > bestScore) {
      bestScore = score;
      bestIndex = i;
    }
  }

  return bestIndex > 0 ? bestIndex : null;
};

const formatCaption = (rawText: string): string => {
  const text = rawText.replace(/\\n/g, '\n').trim();
  if (text.includes('\n') || text.length <= 14) {
    return text;
  }

  const breakIndex = chooseBreakIndex(text);
  if (!breakIndex) {
    return text;
  }

  return `${text.slice(0, breakIndex)}\n${text.slice(breakIndex)}`;
};

export const CaptionStill: React.FC<CaptionStillProps> = ({imageSrc, text}) => {
  const caption = formatCaption(text);

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
          justifyContent: 'flex-start',
          alignItems: 'center',
          padding: '120px 56px 0 56px',
          background: 'linear-gradient(180deg, rgba(0,0,0,0.52) 0%, rgba(0,0,0,0.20) 38%, rgba(0,0,0,0.04) 72%)',
        }}
      >
        <div
          style={{
            width: '100%',
            maxWidth: 930,
            padding: '30px 36px 32px 36px',
            borderRadius: 34,
            background:
              'linear-gradient(140deg, rgba(255, 181, 205, 0.30) 0%, rgba(173, 226, 255, 0.26) 100%)',
            border: '1px solid rgba(255, 255, 255, 0.55)',
            boxShadow: '0 12px 34px rgba(0, 0, 0, 0.22)',
            color: '#ffffff',
            fontFamily: `'${fontFamily}', 'Hiragino Maru Gothic ProN', 'Yu Gothic UI', 'Meiryo', sans-serif`,
            fontSize: 60,
            fontWeight: 400,
            lineHeight: 1.34,
            letterSpacing: '0.03em',
            textAlign: 'center',
            whiteSpace: 'pre-wrap',
            textShadow: '0 6px 16px rgba(0, 0, 0, 0.42)',
            WebkitTextStroke: '1px rgba(18, 18, 18, 0.10)',
          }}
        >
          {caption}
        </div>
      </AbsoluteFill>
    </AbsoluteFill>
  );
};
