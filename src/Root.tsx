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
        }}
      />
    </>
  );
};
