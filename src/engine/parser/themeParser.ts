import { ColorScheme, FontScheme, ResolvedTheme } from '../../types/pptx';
import { DEFAULT_THEME_COLORS, applyColorModifiers } from '../utils/color';
import { XmlUtils } from './xmlUtils';

export class ThemeParser {
  static parse(themeXmlText?: string): ResolvedTheme {
    const colorScheme: ColorScheme = { ...DEFAULT_THEME_COLORS };
    const fontScheme: FontScheme = {
      majorFont: 'Calibri, Arial, sans-serif',
      minorFont: 'Calibri, Arial, sans-serif',
    };

    if (!themeXmlText) {
      return { id: 'default-theme', name: 'Default Theme', colorScheme, fontScheme };
    }

    try {
      const doc = XmlUtils.parseXml(themeXmlText);
      const clrSchemeNode = XmlUtils.findFirstTag(doc, 'clrScheme');

      if (clrSchemeNode) {
        const children = clrSchemeNode.children;
        for (let i = 0; i < children.length; i++) {
          const child = children[i];
          const colorName = child.localName || child.tagName;

          // Check for srgbClr or sysClr
          const srgbNode = XmlUtils.findFirstTag(child, 'srgbClr');
          const sysNode = XmlUtils.findFirstTag(child, 'sysClr');

          let hex = '';
          if (srgbNode) {
            hex = `#${XmlUtils.getAttr(srgbNode, 'val') || '000000'}`;
          } else if (sysNode) {
            hex = `#${XmlUtils.getAttr(sysNode, 'lastClr') || '000000'}`;
          }

          if (hex && colorName) {
            colorScheme[colorName] = hex;
          }
        }
      }

      // Font Scheme
      const fontSchemeNode = XmlUtils.findFirstTag(doc, 'fontScheme');
      if (fontSchemeNode) {
        const majorNode = XmlUtils.findFirstTag(fontSchemeNode, 'majorFont');
        const minorNode = XmlUtils.findFirstTag(fontSchemeNode, 'minorFont');

        if (majorNode) {
          const latinNode = XmlUtils.findFirstTag(majorNode, 'latin');
          if (latinNode) {
            fontScheme.majorFont = XmlUtils.getAttr(latinNode, 'typeface') || fontScheme.majorFont;
          }
        }

        if (minorNode) {
          const latinNode = XmlUtils.findFirstTag(minorNode, 'latin');
          if (latinNode) {
            fontScheme.minorFont = XmlUtils.getAttr(latinNode, 'typeface') || fontScheme.minorFont;
          }
        }
      }
    } catch {
      // Fallback on error
    }

    return {
      id: 'theme-1',
      name: 'OOXML Theme',
      colorScheme,
      fontScheme,
    };
  }
}
