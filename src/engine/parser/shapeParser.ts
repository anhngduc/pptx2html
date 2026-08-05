import {
  BoundingBox,
  ColorScheme,
  FillStyle,
  LineStyle,
  ShapeGeometry,
  Transform2D,
} from '../../types/pptx';
import { applyColorModifiers, resolveSchemeColor } from '../utils/color';
import { emuToPx } from '../utils/emu';
import { XmlUtils } from './xmlUtils';

export class ShapeParser {
  static parseTransform(xfrmNode: Element | null): { bounds: BoundingBox; transform: Transform2D } {
    const bounds: BoundingBox = { xEmu: 0, yEmu: 0, widthEmu: 914400, heightEmu: 914400 };
    const transform: Transform2D = {
      rotationDegrees: 0,
      flipHorizontal: false,
      flipVertical: false,
    };

    if (!xfrmNode) return { bounds, transform };

    const rot = XmlUtils.getIntAttr(xfrmNode, 'rot', 0);
    if (rot) {
      transform.rotationDegrees = rot / 60000;
    }

    transform.flipHorizontal = XmlUtils.getAttr(xfrmNode, 'flipH') === '1';
    transform.flipVertical = XmlUtils.getAttr(xfrmNode, 'flipV') === '1';

    const off = XmlUtils.findFirstTag(xfrmNode, 'off');
    if (off) {
      bounds.xEmu = XmlUtils.getIntAttr(off, 'x', 0);
      bounds.yEmu = XmlUtils.getIntAttr(off, 'y', 0);
    }

    const ext = XmlUtils.findFirstTag(xfrmNode, 'ext');
    if (ext) {
      bounds.widthEmu = XmlUtils.getIntAttr(ext, 'cx', 914400);
      bounds.heightEmu = XmlUtils.getIntAttr(ext, 'cy', 914400);
    }

    return { bounds, transform };
  }

  static parseFill(node: Element | null, themeColors: ColorScheme): FillStyle {
    if (!node) return { type: 'none' };

    const noFill = XmlUtils.findFirstTag(node, 'noFill');
    if (noFill) return { type: 'none' };

    const solidFill = XmlUtils.findFirstTag(node, 'solidFill');
    if (solidFill) {
      const color = ShapeParser.parseColorNode(solidFill, themeColors) || '#3B82F6';
      return { type: 'solid', color, opacity: 1 };
    }

    const gradFill = XmlUtils.findFirstTag(node, 'gradFill');
    if (gradFill) {
      const gsLst = XmlUtils.findFirstTag(gradFill, 'gsLst');
      const stops: { position: number; color: string }[] = [];

      if (gsLst) {
        const gsNodes = XmlUtils.getDirectChildren(gsLst, 'gs');
        for (let i = 0; i < gsNodes.length; i++) {
          const gs = gsNodes[i];
          const pos = XmlUtils.getIntAttr(gs, 'pos', 0) / 100000;
          const color = ShapeParser.parseColorNode(gs, themeColors) || '#2563EB';
          stops.push({ position: pos, color });
        }
      }

      const lin = XmlUtils.findFirstTag(gradFill, 'lin');
      const ang = lin ? XmlUtils.getIntAttr(lin, 'ang', 5400000) / 60000 : 90;

      return {
        type: 'gradient',
        angleDegrees: ang,
        stops: stops.length > 0 ? stops : [{ position: 0, color: '#2563EB' }, { position: 1, color: '#1E40AF' }],
      };
    }

    return { type: 'none' };
  }

  static parseLine(lnNode: Element | null, themeColors: ColorScheme): LineStyle {
    if (!lnNode) {
      return { color: 'transparent', widthPx: 0, dash: 'none' };
    }

    const wEmu = XmlUtils.getIntAttr(lnNode, 'w', 12700);
    const widthPx = Math.max(1, Math.round(emuToPx(wEmu)));

    const solidFill = XmlUtils.findFirstTag(lnNode, 'solidFill');
    const color = solidFill ? ShapeParser.parseColorNode(solidFill, themeColors) || '#9CA3AF' : '#9CA3AF';

    const prstDash = XmlUtils.findFirstTag(lnNode, 'prstDash');
    let dash: LineStyle['dash'] = 'solid';
    if (prstDash) {
      const val = XmlUtils.getAttr(prstDash, 'val');
      if (val === 'dash' || val === 'sysDash') dash = 'dashed';
      if (val === 'dot' || val === 'sysDot') dash = 'dotted';
    }

    const headEnd = XmlUtils.findFirstTag(lnNode, 'headEnd');
    const tailEnd = XmlUtils.findFirstTag(lnNode, 'tailEnd');

    return {
      color,
      widthPx,
      dash,
      startArrow: !!headEnd && XmlUtils.getAttr(headEnd, 'type') !== 'none',
      endArrow: !!tailEnd && XmlUtils.getAttr(tailEnd, 'type') !== 'none',
    };
  }

  static parseGeometry(spPrNode: Element | null): ShapeGeometry {
    if (!spPrNode) return { kind: 'preset', presetName: 'rect' };

    const prstGeom = XmlUtils.findFirstTag(spPrNode, 'prstGeom');
    if (prstGeom) {
      const prst = XmlUtils.getAttr(prstGeom, 'prst') || 'rect';
      return { kind: 'preset', presetName: prst };
    }

    const custGeom = XmlUtils.findFirstTag(spPrNode, 'custGeom');
    if (custGeom) {
      const pathLst = XmlUtils.findFirstTag(custGeom, 'pathLst');
      if (pathLst) {
        const pathNodes = XmlUtils.getDirectChildren(pathLst, 'path');
        let d = '';
        for (let i = 0; i < pathNodes.length; i++) {
          const pathNode = pathNodes[i];
          const children = pathNode.children;
          for (let j = 0; j < children.length; j++) {
            const child = children[j];
            const name = child.localName || child.tagName;
            const pt = XmlUtils.findFirstTag(child, 'pt');
            const x = pt ? emuToPx(XmlUtils.getIntAttr(pt, 'x', 0)) : 0;
            const y = pt ? emuToPx(XmlUtils.getIntAttr(pt, 'y', 0)) : 0;

            if (name === 'moveTo') d += `M ${x} ${y} `;
            if (name === 'lnTo') d += `L ${x} ${y} `;
            if (name === 'close') d += `Z `;
          }
        }
        return { kind: 'custom', svgPath: d.trim() };
      }
    }

    return { kind: 'preset', presetName: 'rect' };
  }

  static parseColorNode(parent: Element | null, themeColors: ColorScheme): string | undefined {
    if (!parent) return undefined;

    const srgbClr = XmlUtils.findFirstTag(parent, 'srgbClr');
    const schemeClr = XmlUtils.findFirstTag(parent, 'schemeClr');
    const sysClr = XmlUtils.findFirstTag(parent, 'sysClr');

    let baseColor = '#000000';
    let targetNode: Element | null = null;

    if (srgbClr) {
      baseColor = `#${XmlUtils.getAttr(srgbClr, 'val') || '000000'}`;
      targetNode = srgbClr;
    } else if (schemeClr) {
      const val = XmlUtils.getAttr(schemeClr, 'val') || 'dk1';
      baseColor = resolveSchemeColor(val, themeColors);
      targetNode = schemeClr;
    } else if (sysClr) {
      baseColor = `#${XmlUtils.getAttr(sysClr, 'lastClr') || '000000'}`;
      targetNode = sysClr;
    } else {
      return undefined;
    }

    if (targetNode) {
      const lumMod = XmlUtils.findFirstTag(targetNode, 'lumMod');
      const lumOff = XmlUtils.findFirstTag(targetNode, 'lumOff');
      const tint = XmlUtils.findFirstTag(targetNode, 'tint');
      const shade = XmlUtils.findFirstTag(targetNode, 'shade');
      const alpha = XmlUtils.findFirstTag(targetNode, 'alpha');

      return applyColorModifiers(baseColor, {
        lumMod: lumMod ? XmlUtils.getIntAttr(lumMod, 'val') : undefined,
        lumOff: lumOff ? XmlUtils.getIntAttr(lumOff, 'val') : undefined,
        tint: tint ? XmlUtils.getIntAttr(tint, 'val') : undefined,
        shade: shade ? XmlUtils.getIntAttr(shade, 'val') : undefined,
        alpha: alpha ? XmlUtils.getIntAttr(alpha, 'val') : undefined,
      });
    }

    return baseColor;
  }
}
