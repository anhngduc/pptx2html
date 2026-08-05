import React from 'react';
import { TextElement } from '../../types/pptx';

interface Props {
  element: TextElement;
}

export const TextElementView: React.FC<Props> = ({ element }) => {
  const { bodyStyle, paragraphs } = element;

  const flexAlignMap = {
    top: 'flex-start',
    middle: 'center',
    bottom: 'flex-end',
  };

  return (
    <div
      className="w-full h-full flex flex-col pointer-events-auto selection:bg-blue-500 selection:text-white"
      style={{
        paddingTop: `${bodyStyle.paddingPx.top}px`,
        paddingRight: `${bodyStyle.paddingPx.right}px`,
        paddingBottom: `${bodyStyle.paddingPx.bottom}px`,
        paddingLeft: `${bodyStyle.paddingPx.left}px`,
        justifyContent: flexAlignMap[bodyStyle.verticalAlign] || 'flex-start',
        overflowWrap: 'break-word',
        wordBreak: 'break-word',
      }}
    >
      {paragraphs.map((p) => (
        <div
          key={p.id}
          className="w-full flex items-baseline"
          style={{
            textAlign: p.alignment,
            paddingLeft: p.indentLevel ? `${p.indentLevel * 18}px` : undefined,
            marginTop: p.spaceBeforePt ? `${p.spaceBeforePt}pt` : undefined,
            marginBottom: p.spaceAfterPt ? `${p.spaceAfterPt}pt` : undefined,
            justifyContent:
              p.alignment === 'center'
                ? 'center'
                : p.alignment === 'right'
                ? 'flex-end'
                : 'flex-start',
          }}
        >
          {p.bullet && p.bullet.type === 'char' && (
            <span className="mr-2 select-none text-slate-400 font-bold">{p.bullet.char}</span>
          )}
          {p.bullet && p.bullet.type === 'number' && (
            <span className="mr-2 select-none text-slate-400 font-bold">{p.bullet.startAt || 1}.</span>
          )}

          <div className="flex-1" style={{ textAlign: p.alignment }}>
            {p.runs.map((r, rIdx) => {
              if (r.type === 'break') {
                return <br key={rIdx} />;
              }

              const spanStyle: React.CSSProperties = {
                fontFamily: r.style.fontFamily,
                fontSize: `${r.style.fontSizePt}pt`,
                color: r.style.color,
                fontWeight: r.style.bold ? 'bold' : 'normal',
                fontStyle: r.style.italic ? 'italic' : 'normal',
                textDecoration: `${r.style.underline ? 'underline ' : ''}${r.style.strikethrough ? 'line-through' : ''}`.trim() || 'none',
                verticalAlign: r.style.superscript ? 'super' : r.style.subscript ? 'sub' : 'baseline',
                lineHeight: 1.35,
              };

              return (
                <span key={rIdx} style={spanStyle}>
                  {r.text}
                </span>
              );
            })}
          </div>
        </div>
      ))}
    </div>
  );
};
