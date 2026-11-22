import React from 'react';

export const SunlitGalleryBackdrop: React.FC = () => {
  return (
    <div
      aria-hidden="true"
      className="sunlit-backdrop pointer-events-none fixed inset-0 overflow-hidden -z-10 select-none"
    >
      {/* Pre-blended daylight/ambient beam sweep (zero runtime blur filter cost) */}
      <div className="sunlit-beam absolute -top-24 -right-20 w-[75vw] h-[70vh] rounded-full opacity-80" />

      {/* Pre-blended soft architectural window shadow */}
      <div className="sunlit-shadow absolute -bottom-24 -left-16 w-[65vw] h-[60vh] rounded-full opacity-55" />

      {/* Soft diagonal botanical shadows using radial gradients (no feGaussianBlur) */}
      <svg
        viewBox="0 0 1440 900"
        preserveAspectRatio="xMidYMid slice"
        className="sunlit-leaves absolute inset-0 w-full h-full opacity-[0.06]"
      >
        <defs>
          <radialGradient id="leafShadowGrad" cx="50%" cy="50%" r="50%">
            <stop offset="0%" stopColor="#111318" stopOpacity="0.95" />
            <stop offset="55%" stopColor="#111318" stopOpacity="0.4" />
            <stop offset="100%" stopColor="#111318" stopOpacity="0" />
          </radialGradient>
        </defs>
        <g fill="url(#leafShadowGrad)">
          <ellipse cx="320" cy="680" rx="250" ry="52" transform="rotate(-28 320 680)" />
          <ellipse cx="470" cy="740" rx="210" ry="44" transform="rotate(-22 470 740)" />
          <ellipse cx="1090" cy="210" rx="280" ry="58" transform="rotate(-34 1090 210)" />
          <ellipse cx="1220" cy="310" rx="230" ry="46" transform="rotate(-26 1220 310)" />
          <ellipse cx="940" cy="540" rx="320" ry="50" transform="rotate(-30 940 540)" />
        </g>
      </svg>
    </div>
  );
};
