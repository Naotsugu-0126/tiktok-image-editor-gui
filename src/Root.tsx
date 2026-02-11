import React from 'react';
import {Composition} from 'remotion';
import {CaptionStill} from './CaptionStill';

export const RemotionRoot: React.FC = () => {
  return (
    <>
      <Composition
        id="CaptionStill"
        component={CaptionStill}
        durationInFrames={1}
        fps={30}
        width={1080}
        height={1920}
        defaultProps={{
          imageSrc: '',
          text: '',
          theme: 'clean',
          font: '標準',
          color: '#FFFFFF',
          strokeColor: '#121212',
          strokeWidth: 0.8,
          telopType: '標準',
          textPosition: '中央',
          previewMode: 'normal',
        }}
      />
    </>
  );
};
