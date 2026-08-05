import { ColorScheme, TableColumn, TableElement, TableRow } from '../../types/pptx';
import { emuToPx } from '../utils/emu';
import { ShapeParser } from './shapeParser';
import { TextParser } from './textParser';
import { XmlUtils } from './xmlUtils';

export class TableParser {
  static parse(
    graphicFrameNode: Element,
    themeColors: ColorScheme,
    majorFont: string
  ): Omit<TableElement, 'id' | 'bounds' | 'transform' | 'zIndex' | 'hidden' | 'opacity' | 'source'> | null {
    const tblNode = XmlUtils.findFirstTag(graphicFrameNode, 'tbl');
    if (!tblNode) return null;

    // Table Grid Columns
    const columns: TableColumn[] = [];
    const tblGrid = XmlUtils.findFirstTag(tblNode, 'tblGrid');
    if (tblGrid) {
      const gridColNodes = XmlUtils.getDirectChildren(tblGrid, 'gridCol');
      for (let i = 0; i < gridColNodes.length; i++) {
        const wEmu = XmlUtils.getIntAttr(gridColNodes[i], 'w', 1828800);
        columns.push({ widthEmu: wEmu });
      }
    }

    // Table Rows
    const rows: TableRow[] = [];
    const trNodes = XmlUtils.getDirectChildren(tblNode, 'tr');

    for (let r = 0; r < trNodes.length; r++) {
      const trNode = trNodes[r];
      const hEmu = XmlUtils.getIntAttr(trNode, 'h', 457200);
      const tcNodes = XmlUtils.getDirectChildren(trNode, 'tc');

      const cells = [];
      for (let c = 0; c < tcNodes.length; c++) {
        const tcNode = tcNodes[c];
        const gridSpan = XmlUtils.getIntAttr(tcNode, 'gridSpan', 1);
        const rowSpan = XmlUtils.getIntAttr(tcNode, 'rowSpan', 1);

        const hMerge = XmlUtils.getAttr(tcNode, 'hMerge') === '1';
        const vMerge = XmlUtils.getAttr(tcNode, 'vMerge') === '1';

        const tcPrNode = XmlUtils.findFirstTag(tcNode, 'tcPr');
        const fill = ShapeParser.parseFill(tcPrNode, themeColors);

        // Borders
        const lnL = XmlUtils.findFirstTag(tcPrNode, 'lnL');
        const lnR = XmlUtils.findFirstTag(tcPrNode, 'lnR');
        const lnT = XmlUtils.findFirstTag(tcPrNode, 'lnT');
        const lnB = XmlUtils.findFirstTag(tcPrNode, 'lnB');

        const borders = {
          left: ShapeParser.parseLine(lnL, themeColors),
          right: ShapeParser.parseLine(lnR, themeColors),
          top: ShapeParser.parseLine(lnT, themeColors),
          bottom: ShapeParser.parseLine(lnB, themeColors),
        };

        const txBodyNode = XmlUtils.findFirstTag(tcNode, 'txBody');
        const textParsed = TextParser.parseTextBody(txBodyNode, themeColors, majorFont);

        cells.push({
          id: `cell-${r}-${c}`,
          rowSpan: rowSpan || 1,
          colSpan: gridSpan || 1,
          fill,
          borders,
          paddingPx: textParsed.bodyStyle.paddingPx,
          verticalAlign: textParsed.bodyStyle.verticalAlign,
          text: {
            paragraphs: textParsed.paragraphs,
          },
          hidden: hMerge || vMerge,
        });
      }

      rows.push({
        heightEmu: hEmu,
        cells,
      });
    }

    return {
      type: 'table',
      columns: columns.length > 0 ? columns : [{ widthEmu: 3657600 }, { widthEmu: 3657600 }],
      rows,
    };
  }
}
