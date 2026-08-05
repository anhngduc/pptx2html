import React from 'react';
import { SlideElement } from '../types/pptx';
import { emuToPx } from '../engine/utils/emu';
import { TextElementView } from './renderers/TextElementView';
import { ShapeElementView } from './renderers/ShapeElementView';
import { TableElementView } from './renderers/TableElementView';
import { ChartElementView } from './renderers/ChartElementView';
import {
  ConnectorElementView,
  ImageElementView,
  MediaElementView,
  OleElementView,
  UnknownElementView,
} from './renderers/MediaElementView';

interface Props {
  element: SlideElement;
  isSelected: boolean;
  isHovered: boolean;
  onSelect: (element: SlideElement) => void;
  onHover: (element: SlideElement | null) => void;
  scaleRatio: number;
}

export const ElementRenderer: React.FC<Props> = ({
  element,
  isSelected,
  isHovered,
  onSelect,
  onHover,
}) => {
  if (element.hidden) return null;

  const leftPx = emuToPx(element.bounds.xEmu);
  const topPx = emuToPx(element.bounds.yEmu);
  const widthPx = emuToPx(element.bounds.widthEmu);
  const heightPx = emuToPx(element.bounds.heightEmu);

  const style: React.CSSProperties = {
    position: 'absolute',
    left: `${leftPx}px`,
    top: `${topPx}px`,
    width: `${widthPx}px`,
    height: `${heightPx}px`,
    transform: `rotate(${element.transform.rotationDegrees}deg)${
      element.transform.flipHorizontal ? ' scaleX(-1)' : ''
    }${element.transform.flipVertical ? ' scaleY(-1)' : ''}`,
    zIndex: element.zIndex,
    opacity: element.opacity,
  };

  const renderContent = () => {
    switch (element.type) {
      case 'text':
        return <TextElementView element={element} />;
      case 'shape':
        return <ShapeElementView element={element} widthPx={widthPx} heightPx={heightPx} />;
      case 'table':
        return <TableElementView element={element} />;
      case 'chart':
        return <ChartElementView element={element} />;
      case 'image':
        return <ImageElementView element={element} />;
      case 'media':
        return <MediaElementView element={element} />;
      case 'ole':
        return <OleElementView element={element} />;
      case 'connector':
        return <ConnectorElementView element={element} />;
      case 'unknown':
        return <UnknownElementView element={element} />;
      default:
        return null;
    }
  };

  return (
    <div
      style={style}
      onClick={(e) => {
        e.stopPropagation();
        onSelect(element);
      }}
      onMouseEnter={() => onHover(element)}
      onMouseLeave={() => onHover(null)}
      className={`group cursor-pointer transition-all duration-150 ${
        isSelected
          ? 'ring-2 ring-blue-500 ring-offset-2 ring-offset-slate-900 rounded-sm'
          : isHovered
          ? 'ring-1 ring-blue-400/80 ring-offset-1 ring-offset-slate-900 rounded-sm'
          : ''
      }`}
    >
      {renderContent()}

      {/* Selection Bounding Box Indicator */}
      {(isSelected || isHovered) && (
        <div className="absolute -top-6 left-0 bg-blue-600 text-white text-[10px] font-mono px-1.5 py-0.5 rounded shadow pointer-events-none z-50 whitespace-nowrap">
          {element.name || element.type} ({Math.round(widthPx)}x{Math.round(heightPx)}px)
        </div>
      )}
    </div>
  );
};
