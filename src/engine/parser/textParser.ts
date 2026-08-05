import {
  BulletStyle,
  ColorScheme,
  TextAlignment,
  TextBodyStyle,
  TextParagraph,
  TextRun,
  TextRunStyle,
} from '../../types/pptx';
import { applyColorModifiers, resolveSchemeColor } from '../utils/color';
import { hundredthPtToPt } from '../utils/emu';
import { XmlUtils } from './xmlUtils';

export class TextParser {
  static parseTextBody(txBodyNode: Element | null, themeColors: ColorScheme, majorFont: string): {
    bodyStyle: TextBodyStyle;
    paragraphs: TextParagraph[];
  } {
    const bodyStyle: TextBodyStyle = {
      paddingPx: { top: 8, right: 10, bottom: 8, left: 10 },
      verticalAlign: 'top',
      wrapText: true,
      autofit: 'none',
    };

    const paragraphs: TextParagraph[] = [];

    if (!txBodyNode) {
      return { bodyStyle, paragraphs };
    }

    // Body Pr
    const bodyPr = XmlUtils.findFirstTag(txBodyNode, 'bodyPr');
    if (bodyPr) {
      const anchor = XmlUtils.getAttr(bodyPr, 'anchor');
      if (anchor === 'ctr') bodyStyle.verticalAlign = 'middle';
      if (anchor === 'b') bodyStyle.verticalAlign = 'bottom';

      const wrap = XmlUtils.getAttr(bodyPr, 'wrap');
      if (wrap === 'none') bodyStyle.wrapText = false;

      const lIns = XmlUtils.getIntAttr(bodyPr, 'lIns', 91440);
      const tIns = XmlUtils.getIntAttr(bodyPr, 'tIns', 45720);
      const rIns = XmlUtils.getIntAttr(bodyPr, 'rIns', 91440);
      const bIns = XmlUtils.getIntAttr(bodyPr, 'bIns', 45720);

      bodyStyle.paddingPx = {
        top: Math.round((tIns / 914400) * 96),
        right: Math.round((rIns / 914400) * 96),
        bottom: Math.round((bIns / 914400) * 96),
        left: Math.round((lIns / 914400) * 96),
      };
    }

    // Paragraphs
    const pNodes = XmlUtils.getDirectChildren(txBodyNode, 'p');
    for (let i = 0; i < pNodes.length; i++) {
      const pNode = pNodes[i];
      const pPrNode = XmlUtils.findFirstTag(pNode, 'pPr');

      let alignment: TextAlignment = 'left';
      let indentLevel = 0;
      let bullet: BulletStyle | undefined = undefined;
      let spaceBeforePt: number | undefined;
      let spaceAfterPt: number | undefined;

      if (pPrNode) {
        const algn = XmlUtils.getAttr(pPrNode, 'algn');
        if (algn === 'ctr') alignment = 'center';
        if (algn === 'r') alignment = 'right';
        if (algn === 'just') alignment = 'justify';

        indentLevel = XmlUtils.getIntAttr(pPrNode, 'lvl', 0);

        // Bullets
        const buChar = XmlUtils.findFirstTag(pPrNode, 'buChar');
        const buAutoNum = XmlUtils.findFirstTag(pPrNode, 'buAutoNum');
        const buNone = XmlUtils.findFirstTag(pPrNode, 'buNone');

        if (buChar) {
          bullet = { type: 'char', char: XmlUtils.getAttr(buChar, 'char') || '•' };
        } else if (buAutoNum) {
          bullet = { type: 'number', startAt: XmlUtils.getIntAttr(buAutoNum, 'startAt', 1) };
        } else if (buNone) {
          bullet = { type: 'none' };
        } else if (indentLevel > 0) {
          bullet = { type: 'char', char: indentLevel % 2 === 1 ? '–' : '•' };
        }

        // Spacing
        const spcBef = XmlUtils.findFirstTag(pPrNode, 'spcBef');
        if (spcBef) {
          const spcPts = XmlUtils.findFirstTag(spcBef, 'spcPts');
          if (spcPts) spaceBeforePt = hundredthPtToPt(XmlUtils.getIntAttr(spcPts, 'val', 0));
        }

        const spcAft = XmlUtils.findFirstTag(pPrNode, 'spcAft');
        if (spcAft) {
          const spcPts = XmlUtils.findFirstTag(spcAft, 'spcPts');
          if (spcPts) spaceAfterPt = hundredthPtToPt(XmlUtils.getIntAttr(spcPts, 'val', 0));
        }
      }

      // Runs
      const runs: TextRun[] = [];
      const children = pNode.children;

      for (let j = 0; j < children.length; j++) {
        const child = children[j];
        const localName = child.localName || child.tagName;

        if (localName === 'r') {
          const tNode = XmlUtils.findFirstTag(child, 't');
          const rPrNode = XmlUtils.findFirstTag(child, 'rPr');
          const text = tNode ? tNode.textContent || '' : '';

          const runStyle = TextParser.parseRunStyle(rPrNode, themeColors, majorFont);
          runs.push({
            type: 'text',
            text,
            style: runStyle,
          });
        } else if (localName === 'br') {
          runs.push({
            type: 'break',
            text: '\n',
            style: TextParser.parseRunStyle(null, themeColors, majorFont),
          });
        } else if (localName === 'fld') {
          const tNode = XmlUtils.findFirstTag(child, 't');
          const text = tNode ? tNode.textContent || '' : '';
          const rPrNode = XmlUtils.findFirstTag(child, 'rPr');
          runs.push({
            type: 'field',
            text,
            fieldType: XmlUtils.getAttr(child, 'type') || 'slideNum',
            style: TextParser.parseRunStyle(rPrNode, themeColors, majorFont),
          });
        }
      }

      if (runs.length === 0) {
        // Empty paragraph line break
        runs.push({
          type: 'text',
          text: '',
          style: TextParser.parseRunStyle(null, themeColors, majorFont),
        });
      }

      paragraphs.push({
        id: `p-${i}`,
        alignment,
        indentLevel,
        spaceBeforePt,
        spaceAfterPt,
        bullet,
        runs,
      });
    }

    return { bodyStyle, paragraphs };
  }

  static parseRunStyle(rPrNode: Element | null, themeColors: ColorScheme, defaultFont: string): TextRunStyle {
    const style: TextRunStyle = {
      fontFamily: defaultFont,
      fontSizePt: 14,
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      color: themeColors.dk1 || '#000000',
    };

    if (!rPrNode) return style;

    const sz = XmlUtils.getIntAttr(rPrNode, 'sz', 0);
    if (sz > 0) {
      style.fontSizePt = hundredthPtToPt(sz);
    }

    style.bold = XmlUtils.getAttr(rPrNode, 'b') === '1';
    style.italic = XmlUtils.getAttr(rPrNode, 'i') === '1';
    style.underline = XmlUtils.getAttr(rPrNode, 'u') === 'sng';
    style.strikethrough = XmlUtils.getAttr(rPrNode, 'strike') === 'sng';

    const latin = XmlUtils.findFirstTag(rPrNode, 'latin');
    if (latin) {
      const typeface = XmlUtils.getAttr(latin, 'typeface');
      if (typeface) style.fontFamily = typeface;
    }

    // Color
    const solidFill = XmlUtils.findFirstTag(rPrNode, 'solidFill');
    if (solidFill) {
      const srgbClr = XmlUtils.findFirstTag(solidFill, 'srgbClr');
      const schemeClr = XmlUtils.findFirstTag(solidFill, 'schemeClr');

      if (srgbClr) {
        style.color = `#${XmlUtils.getAttr(srgbClr, 'val') || '000000'}`;
      } else if (schemeClr) {
        const val = XmlUtils.getAttr(schemeClr, 'val') || 'dk1';
        style.color = resolveSchemeColor(val, themeColors);
      }
    }

    return style;
  }
}
