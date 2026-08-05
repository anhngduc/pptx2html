import React, { useRef, useState } from 'react';
import { SlideDocument, SlideElement, SlideViewport as ViewportType } from '../types/pptx';
import { emuToPx } from '../engine/utils/emu';
import { ElementRenderer } from './ElementRenderer';
import { Maximize2, Minimize2, MousePointer, ZoomIn, ZoomOut } from 'lucide-react';

interface Props {
  slide: SlideDocument;
  viewport: ViewportType;
  selectedElement: SlideElement | null;
  onSelectElement: (element: SlideElement | null) => void;
  scaleMode: 'fit' | '100' | '125' | '150';
  onScaleChange: (mode: 'fit' | '100' | '125' | '150') => void;
}

export const SlideViewport: React.FC<Props> = ({
  slide,
  viewport,
  selectedElement,
  onSelectElement,
  scaleMode,
  onScaleChange,
}) => {
  const containerRef = useRef<HTMLDivElement>(null);
  const [hoveredElement, setHoveredElement] = useState<SlideElement | null>(null);
  const [laserActive, setLaserActive] = useState(false);
  const [laserPos, setLaserPos] = useState({ x: 0, y: 0 });

  const slideWidthPx = Math.round(emuToPx(viewport.widthEmu));
  const slideHeightPx = Math.round(emuToPx(viewport.heightEmu));

  // Compute CSS background
  const getSlideBackgroundStyle = (): React.CSSProperties => {
    const bg = slide.background;
    if (bg.type === 'solid') {
      return { backgroundColor: bg.color };
    }
    if (bg.type === 'gradient') {
      const stops = bg.stops.map((s) => `${s.color} ${Math.round(s.position * 100)}%`).join(', ');
      return { background: `linear-gradient(${bg.angleDegrees}deg, ${stops})` };
    }
    return { backgroundColor: '#FFFFFF' };
  };

  const handleMouseMove = (e: React.MouseEvent<HTMLDivElement>) => {
    if (!laserActive || !containerRef.current) return;
    const rect = containerRef.current.getBoundingClientRect();
    setLaserPos({
      x: e.clientX - rect.left,
      y: e.clientY - rect.top,
    });
  };

  return (
    <div
      ref={containerRef}
      className="relative w-full h-full bg-[#D9D7D1] flex flex-col items-center justify-center overflow-auto p-4 select-none"
      onClick={() => onSelectElement(null)}
      onMouseMove={handleMouseMove}
    >
      {/* Viewport Control Bar */}
      <div className="absolute top-4 right-4 z-40 bg-[#E4E3E0] border border-[#141414] p-1 flex items-center gap-1.5 font-mono text-xs shadow-md">
        <button
          onClick={(e) => {
            e.stopPropagation();
            setLaserActive(!laserActive);
          }}
          className={`px-2 py-1 text-xs flex items-center gap-1 transition-colors border ${
            laserActive ? 'bg-red-700 text-white border-[#141414]' : 'bg-[#E4E3E0] text-[#141414] border-transparent hover:border-[#141414]'
          }`}
          title="Laser Pointer Tool"
        >
          <MousePointer className="w-3.5 h-3.5" />
          <span className="hidden sm:inline">LASER</span>
        </button>

        <div className="h-4 w-[1px] bg-[#141414]/30" />

        <button
          onClick={(e) => {
            e.stopPropagation();
            onScaleChange('fit');
          }}
          className={`px-2 py-1 text-xs font-mono transition-colors border ${
            scaleMode === 'fit' ? 'bg-[#141414] text-[#E4E3E0] border-[#141414]' : 'text-[#141414] border-transparent hover:border-[#141414]'
          }`}
        >
          FIT
        </button>
        <button
          onClick={(e) => {
            e.stopPropagation();
            onScaleChange('100');
          }}
          className={`px-2 py-1 text-xs font-mono transition-colors border ${
            scaleMode === '100' ? 'bg-[#141414] text-[#E4E3E0] border-[#141414]' : 'text-[#141414] border-transparent hover:border-[#141414]'
          }`}
        >
          100%
        </button>
        <button
          onClick={(e) => {
            e.stopPropagation();
            onScaleChange('125');
          }}
          className={`px-2 py-1 text-xs font-mono transition-colors border ${
            scaleMode === '125' ? 'bg-[#141414] text-[#E4E3E0] border-[#141414]' : 'text-[#141414] border-transparent hover:border-[#141414]'
          }`}
        >
          125%
        </button>
      </div>

      {/* Main Slide Canvas Container */}
      <div
        className="relative shadow-2xl overflow-hidden border-2 border-[#141414] transition-all duration-200"
        style={{
          width: `${slideWidthPx}px`,
          height: `${slideHeightPx}px`,
          transform: scaleMode === 'fit' ? 'scale(0.85)' : scaleMode === '125' ? 'scale(1.1)' : 'scale(1)',
          transformOrigin: 'center center',
          ...getSlideBackgroundStyle(),
        }}
      >
        {/* Render Slide Elements */}
        {slide.elements.map((element) => (
          <ElementRenderer
            key={element.id}
            element={element}
            isSelected={selectedElement?.id === element.id}
            isHovered={hoveredElement?.id === element.id}
            onSelect={(el) => onSelectElement(el)}
            onHover={(el) => setHoveredElement(el)}
            scaleRatio={1}
          />
        ))}

        {/* Laser Pointer Effect */}
        {laserActive && (
          <div
            className="absolute pointer-events-none z-50 w-5 h-5 rounded-full bg-red-500 shadow-[0_0_15px_rgba(239,68,68,1)] transform -translate-x-1/2 -translate-y-1/2 animate-pulse"
            style={{ left: `${laserPos.x}px`, top: `${laserPos.y}px` }}
          />
        )}
      </div>
    </div>
  );
};
