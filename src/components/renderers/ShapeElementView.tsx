import React from 'react';
import { ShapeElement } from '../../types/pptx';
import { TextElementView } from './TextElementView';

interface Props {
  element: ShapeElement;
  widthPx: number;
  heightPx: number;
}

export const ShapeElementView: React.FC<Props> = ({ element, widthPx, heightPx }) => {
  const { geometry, fill, line, text } = element;
  const gradientId = `grad-${element.id}`;

  // Build SVG path based on preset geometry
  const getSvgPath = () => {
    if (geometry.kind === 'custom' && geometry.svgPath) {
      return geometry.svgPath;
    }

    const preset = geometry.presetName || 'rect';
    const w = widthPx;
    const h = heightPx;

    switch (preset) {
      case 'ellipse':
        return `M ${w / 2} 0 A ${w / 2} ${h / 2} 0 1 1 ${w / 2} ${h} A ${w / 2} ${h / 2} 0 1 1 ${w / 2} 0 Z`;
      case 'roundRect': {
        const r = Math.min(w, h) * 0.15;
        return `M ${r} 0 L ${w - r} 0 A ${r} ${r} 0 0 1 ${w} ${r} L ${w} ${h - r} A ${r} ${r} 0 0 1 ${w - r} ${h} L ${r} ${h} A ${r} ${r} 0 0 1 0 ${h - r} L 0 ${r} A ${r} ${r} 0 0 1 ${r} 0 Z`;
      }
      case 'triangle':
        return `M ${w / 2} 0 L ${w} ${h} L 0 ${h} Z`;
      case 'rightTriangle':
        return `M 0 0 L ${w} ${h} L 0 ${h} Z`;
      case 'diamond':
        return `M ${w / 2} 0 L ${w} ${h / 2} L ${w / 2} ${h} L 0 ${h / 2} Z`;
      case 'chevron':
        return `M 0 0 L ${w * 0.75} 0 L ${w} ${h / 2} L ${w * 0.75} ${h} L 0 ${h} L ${w * 0.25} ${h / 2} Z`;
      case 'star5': {
        const cx = w / 2;
        const cy = h / 2;
        const outerR = Math.min(w, h) / 2;
        const innerR = outerR * 0.4;
        let pts = '';
        for (let i = 0; i < 10; i++) {
          const r = i % 2 === 0 ? outerR : innerR;
          const angle = (i * Math.PI) / 5 - Math.PI / 2;
          const x = cx + r * Math.cos(angle);
          const y = cy + r * Math.sin(angle);
          pts += `${i === 0 ? 'M' : 'L'} ${x} ${y} `;
        }
        return `${pts} Z`;
      }
      default:
        // rect
        return `M 0 0 L ${w} 0 L ${w} ${h} L 0 ${h} Z`;
    }
  };

  const pathD = getSvgPath();

  const fillAttr =
    fill.type === 'solid'
      ? fill.color
      : fill.type === 'gradient'
      ? `url(#${gradientId})`
      : 'none';

  return (
    <div className="relative w-full h-full">
      <svg className="absolute inset-0 w-full h-full overflow-visible pointer-events-none">
        <defs>
          {fill.type === 'gradient' && (
            <linearGradient
              id={gradientId}
              x1="0%"
              y1="0%"
              x2={Math.cos((fill.angleDegrees * Math.PI) / 180) * 100 + '%'}
              y2={Math.sin((fill.angleDegrees * Math.PI) / 180) * 100 + '%'}
            >
              {fill.stops.map((stop, sIdx) => (
                <stop
                  key={sIdx}
                  offset={`${Math.round(stop.position * 100)}%`}
                  stopColor={stop.color}
                />
              ))}
            </linearGradient>
          )}
        </defs>

        <path
          d={pathD}
          fill={fillAttr}
          stroke={line.widthPx > 0 ? line.color : 'none'}
          strokeWidth={line.widthPx}
          strokeDasharray={line.dash === 'dashed' ? '6,6' : line.dash === 'dotted' ? '2,2' : undefined}
        />
      </svg>

      {/* Text Overlay inside shape */}
      {text && (
        <div className="absolute inset-0 z-10 flex items-center justify-center p-2">
          <TextElementView
            element={{
              id: `${element.id}-text`,
              type: 'text',
              bounds: element.bounds,
              transform: element.transform,
              zIndex: 1,
              hidden: false,
              opacity: 1,
              bodyStyle: text.bodyStyle,
              paragraphs: text.paragraphs,
              source: element.source,
            }}
          />
        </div>
      )}
    </div>
  );
};
