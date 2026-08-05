export class XmlUtils {
  static parseXml(xmlText: string): Document {
    const parser = new DOMParser();
    let doc = parser.parseFromString(xmlText, 'text/xml');
    let parserError = doc.querySelector('parsererror');

    if (parserError) {
      // If parsing failed due to missing namespace declarations (e.g. prefix r, p, a, c, id, etc.),
      // auto-inject standard OOXML namespace attributes into the root element and re-parse.
      const ooxmlNamespaces = [
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"',
        'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"',
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
        'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"',
        'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"',
        'xmlns:v="urn:schemas-microsoft-com:vml"',
        'xmlns:id="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
      ].join(' ');

      const patchedXml = xmlText.replace(/(<[a-zA-Z0-9_:-]+)(\s|>|\/>)/, (match, p1, p2) => {
        return `${p1} ${ooxmlNamespaces}${p2}`;
      });

      doc = parser.parseFromString(patchedXml, 'text/xml');
      parserError = doc.querySelector('parsererror');
    }

    if (parserError) {
      throw new Error(`XML Parse Error: ${parserError.textContent}`);
    }
    return doc;
  }

  static getAttr(element: Element | null, attrName: string): string | null {
    if (!element) return null;
    return element.getAttribute(attrName);
  }

  static getIntAttr(element: Element | null, attrName: string, defaultValue = 0): number {
    const val = XmlUtils.getAttr(element, attrName);
    if (!val) return defaultValue;
    const num = parseInt(val, 10);
    return isNaN(num) ? defaultValue : num;
  }

  static getDirectChild(parent: Element | null, tagName: string): Element | null {
    if (!parent) return null;
    for (let i = 0; i < parent.children.length; i++) {
      const child = parent.children[i];
      if (child.localName === tagName || child.tagName === tagName) {
        return child;
      }
    }
    return null;
  }

  static getDirectChildren(parent: Element | null, tagName?: string): Element[] {
    if (!parent) return [];
    const results: Element[] = [];
    for (let i = 0; i < parent.children.length; i++) {
      const child = parent.children[i];
      if (!tagName || child.localName === tagName || child.tagName === tagName) {
        results.push(child);
      }
    }
    return results;
  }

  static findFirstTag(parent: Element | Document | null, tagName: string): Element | null {
    if (!parent) return null;
    const elements = parent.getElementsByTagName('*');
    for (let i = 0; i < elements.length; i++) {
      const el = elements[i];
      if (el.localName === tagName || el.tagName === tagName) {
        return el;
      }
    }
    return null;
  }
}
