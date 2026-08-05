import React from 'react';
import { TableElement } from '../../types/pptx';

interface Props {
  element: TableElement;
}

export const TableElementView: React.FC<Props> = ({ element }) => {
  return (
    <div className="w-full h-full overflow-hidden border-collapse">
      <table className="w-full h-full border-collapse table-fixed">
        <colgroup>
          {element.columns.map((col, idx) => (
            <col key={idx} style={{ width: `${col.widthEmu}emu` }} />
          ))}
        </colgroup>
        <tbody>
          {element.rows.map((row, rIdx) => (
            <tr key={rIdx} style={{ height: `${row.heightEmu}emu` }}>
              {row.cells.map((cell, cIdx) => {
                if (cell.hidden) return null;

                const bgStyle =
                  cell.fill.type === 'solid'
                    ? { backgroundColor: cell.fill.color }
                    : { backgroundColor: 'transparent' };

                const borderTop = cell.borders.top?.widthPx
                  ? `${cell.borders.top.widthPx}px solid ${cell.borders.top.color}`
                  : '1px solid #CBD5E1';
                const borderRight = cell.borders.right?.widthPx
                  ? `${cell.borders.right.widthPx}px solid ${cell.borders.right.color}`
                  : '1px solid #CBD5E1';
                const borderBottom = cell.borders.bottom?.widthPx
                  ? `${cell.borders.bottom.widthPx}px solid ${cell.borders.bottom.color}`
                  : '1px solid #CBD5E1';
                const borderLeft = cell.borders.left?.widthPx
                  ? `${cell.borders.left.widthPx}px solid ${cell.borders.left.color}`
                  : '1px solid #CBD5E1';

                return (
                  <td
                    key={cell.id}
                    rowSpan={cell.rowSpan > 1 ? cell.rowSpan : undefined}
                    colSpan={cell.colSpan > 1 ? cell.colSpan : undefined}
                    style={{
                      ...bgStyle,
                      borderTop,
                      borderRight,
                      borderBottom,
                      borderLeft,
                      paddingTop: `${cell.paddingPx.top}px`,
                      paddingRight: `${cell.paddingPx.right}px`,
                      paddingBottom: `${cell.paddingPx.bottom}px`,
                      paddingLeft: `${cell.paddingPx.left}px`,
                      verticalAlign: cell.verticalAlign || 'middle',
                    }}
                  >
                    {cell.text.paragraphs.map((p, pIdx) => (
                      <div key={pIdx} style={{ textAlign: p.alignment }}>
                        {p.runs.map((r, rIdx) => (
                          <span
                            key={rIdx}
                            style={{
                              fontFamily: r.style.fontFamily,
                              fontSize: `${r.style.fontSizePt}pt`,
                              color: r.style.color,
                              fontWeight: r.style.bold ? 'bold' : 'normal',
                              fontStyle: r.style.italic ? 'italic' : 'normal',
                            }}
                          >
                            {r.text}
                          </span>
                        ))}
                      </div>
                    ))}
                  </td>
                );
              })}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
};
