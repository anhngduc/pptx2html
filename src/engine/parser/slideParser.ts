import {
  Diagnostic,
  PresentationDocument,
  ResourceStore,
  SlideDocument,
  SlideElement,
  SlideViewport,
} from '../../types/pptx';
import { OoxmlPackageReader } from '../package/packageReader';
import { DEFAULT_SLIDE_HEIGHT_EMU, DEFAULT_SLIDE_WIDTH_EMU } from '../utils/emu';
import { ChartParser } from './chartParser';
import { ShapeParser } from './shapeParser';
import { TableParser } from './tableParser';
import { TextParser } from './textParser';
import { ThemeParser } from './themeParser';
import { XmlUtils } from './xmlUtils';

export class SlideParser {
  static async parsePresentation(
    packageReader: OoxmlPackageReader,
    resourceStore: ResourceStore
  ): Promise<PresentationDocument> {
    const diagnostics: Diagnostic[] = [];
    const timestamp = Date.now();

    const addDiag = (severity: Diagnostic['severity'], code: string, message: string, slideIndex?: number) => {
      diagnostics.push({
        id: `diag-${diagnostics.length + 1}`,
        severity,
        code,
        message,
        slideIndex,
        timestamp,
      });
    };

    addDiag('info', 'PKG_LOADED', 'OOXML Package unpacked and indexed successfully.');

    // 1. Presentation.xml
    const presPart = packageReader.getPart('ppt/presentation.xml');
    if (!presPart || !presPart.textData) {
      addDiag('error', 'PKG_MISSING_PRES', 'ppt/presentation.xml not found in OOXML package.');
      throw new Error('Invalid PPTX: ppt/presentation.xml is missing.');
    }

    const presDoc = XmlUtils.parseXml(presPart.textData);

    // Slide Viewport Size
    const viewport: SlideViewport = {
      widthEmu: DEFAULT_SLIDE_WIDTH_EMU,
      heightEmu: DEFAULT_SLIDE_HEIGHT_EMU,
      aspectRatio: 16 / 9,
    };

    const sldSz = XmlUtils.findFirstTag(presDoc, 'sldSz');
    if (sldSz) {
      const cx = XmlUtils.getIntAttr(sldSz, 'cx', DEFAULT_SLIDE_WIDTH_EMU);
      const cy = XmlUtils.getIntAttr(sldSz, 'cy', DEFAULT_SLIDE_HEIGHT_EMU);
      viewport.widthEmu = cx;
      viewport.heightEmu = cy;
      viewport.aspectRatio = cx / cy;
      addDiag('info', 'VIEWPORT_RESOLVED', `Slide Viewport: ${cx}x${cy} EMUs (${(cx / 914400 * 96).toFixed(0)}x${(cy / 914400 * 96).toFixed(0)}px).`);
    }

    // Theme
    const themePart = packageReader.getPart('ppt/theme/theme1.xml');
    const theme = ThemeParser.parse(themePart?.textData);
    addDiag('info', 'THEME_RESOLVED', `Theme resolved: ${theme.name} with ${Object.keys(theme.colorScheme).length} color slots.`);

    // Slide IDs List
    const sldIdLst = XmlUtils.findFirstTag(presDoc, 'sldIdLst');
    const slidePaths: string[] = [];

    if (sldIdLst) {
      const sldIdNodes = XmlUtils.getDirectChildren(sldIdLst, 'sldId');
      for (let i = 0; i < sldIdNodes.length; i++) {
        const rId = XmlUtils.getAttr(sldIdNodes[i], 'id') || XmlUtils.getAttr(sldIdNodes[i], 'r:id') || '';
        const resolved = packageReader.resolveTarget('ppt/presentation.xml', rId);
        if (resolved.path) {
          slidePaths.push(resolved.path);
        }
      }
    }

    // Fallback if sldIdLst failed
    if (slidePaths.length === 0) {
      for (let i = 1; i <= 50; i++) {
        const path = `ppt/slides/slide${i}.xml`;
        if (packageReader.getPart(path)) {
          slidePaths.push(path);
        } else {
          break;
        }
      }
    }

    addDiag('info', 'SLIDES_INDEXED', `Found ${slidePaths.length} slide parts in presentation sequence.`);

    const slides: SlideDocument[] = [];

    // Parse each slide
    for (let index = 0; index < slidePaths.length; index++) {
      const slidePath = slidePaths[index];
      const slidePart = packageReader.getPart(slidePath);
      if (!slidePart || !slidePart.textData) continue;

      try {
        const slideDoc = XmlUtils.parseXml(slidePart.textData);
        const elements: SlideElement[] = [];

        // Background
        const bgNode = XmlUtils.findFirstTag(slideDoc, 'bg');
        const bgPrNode = bgNode ? XmlUtils.findFirstTag(bgNode, 'bgPr') : null;
        const background = bgPrNode
          ? ShapeParser.parseFill(bgPrNode, theme.colorScheme)
          : { type: 'solid' as const, color: theme.colorScheme.lt1 || '#FFFFFF', opacity: 1 };

        // spTree
        const spTree = XmlUtils.findFirstTag(slideDoc, 'spTree');
        if (spTree) {
          const children = spTree.children;
          let zIndexCounter = 1;

          for (let i = 0; i < children.length; i++) {
            const child = children[i];
            const tagName = child.localName || child.tagName;

            // 1. Shapes (p:sp)
            if (tagName === 'sp') {
              const spPr = XmlUtils.findFirstTag(child, 'spPr');
              const { bounds, transform } = ShapeParser.parseTransform(spPr ? XmlUtils.findFirstTag(spPr, 'xfrm') : null);
              const fill = ShapeParser.parseFill(spPr, theme.colorScheme);
              const line = ShapeParser.parseLine(spPr ? XmlUtils.findFirstTag(spPr, 'ln') : null, theme.colorScheme);
              const geometry = ShapeParser.parseGeometry(spPr);

              const txBody = XmlUtils.findFirstTag(child, 'txBody');
              const textParsed = txBody ? TextParser.parseTextBody(txBody, theme.colorScheme, theme.fontScheme.majorFont) : undefined;

              // Check if it's text-only vs shape with text
              const name = XmlUtils.getAttr(XmlUtils.findFirstTag(child, 'cNvPr'), 'name') || `Shape ${i + 1}`;

              if (fill.type === 'none' && line.dash === 'none' && textParsed && textParsed.paragraphs.length > 0) {
                elements.push({
                  id: `el-${slidePath}-${i}`,
                  type: 'text',
                  name,
                  bounds,
                  transform,
                  zIndex: zIndexCounter++,
                  hidden: false,
                  opacity: 1,
                  bodyStyle: textParsed.bodyStyle,
                  paragraphs: textParsed.paragraphs,
                  source: { partPath: slidePath, xmlPath: `spTree/sp[${i}]` },
                });
              } else {
                elements.push({
                  id: `el-${slidePath}-${i}`,
                  type: 'shape',
                  name,
                  bounds,
                  transform,
                  zIndex: zIndexCounter++,
                  hidden: false,
                  opacity: 1,
                  geometry,
                  fill,
                  line,
                  text: textParsed && textParsed.paragraphs.length > 0 ? textParsed : undefined,
                  source: { partPath: slidePath, xmlPath: `spTree/sp[${i}]` },
                });
              }
            }

            // 2. Pictures (p:pic)
            else if (tagName === 'pic') {
              const spPr = XmlUtils.findFirstTag(child, 'spPr');
              const { bounds, transform } = ShapeParser.parseTransform(spPr ? XmlUtils.findFirstTag(spPr, 'xfrm') : null);
              const blip = XmlUtils.findFirstTag(child, 'blip');
              const embedRid = blip ? XmlUtils.getAttr(blip, 'embed') || XmlUtils.getAttr(blip, 'r:embed') : null;

              let srcUrl = '';
              let mimeType = 'image/png';

              if (embedRid) {
                const resolved = packageReader.resolveTarget(slidePath, embedRid);
                if (resolved.path) {
                  const imageResource = resourceStore.images.get(resolved.path);
                  if (imageResource) {
                    srcUrl = imageResource.url;
                    mimeType = imageResource.mimeType;
                  }
                }
              }

              elements.push({
                id: `el-${slidePath}-${i}`,
                type: 'image',
                name: XmlUtils.getAttr(XmlUtils.findFirstTag(child, 'cNvPr'), 'name') || 'Picture',
                bounds,
                transform,
                zIndex: zIndexCounter++,
                hidden: false,
                opacity: 1,
                resourceId: embedRid || 'pic-1',
                mimeType,
                srcUrl,
                source: { partPath: slidePath, xmlPath: `spTree/pic[${i}]` },
              });
            }

            // 3. Graphic Frame (p:graphicFrame - Table, Chart, SmartArt)
            else if (tagName === 'graphicFrame') {
              const xfrm = XmlUtils.findFirstTag(child, 'xfrm');
              const { bounds, transform } = ShapeParser.parseTransform(xfrm);

              // Check if Table
              const tableParsed = TableParser.parse(child, theme.colorScheme, theme.fontScheme.majorFont);
              if (tableParsed) {
                elements.push({
                  ...tableParsed,
                  id: `el-${slidePath}-${i}`,
                  name: XmlUtils.getAttr(XmlUtils.findFirstTag(child, 'cNvPr'), 'name') || 'Table',
                  bounds,
                  transform,
                  zIndex: zIndexCounter++,
                  hidden: false,
                  opacity: 1,
                  source: { partPath: slidePath, xmlPath: `spTree/graphicFrame[${i}]` },
                });
                continue;
              }

              // Check if Chart
              const chartNode = XmlUtils.findFirstTag(child, 'chart');
              if (chartNode) {
                const rId = XmlUtils.getAttr(chartNode, 'id') || XmlUtils.getAttr(chartNode, 'r:id');
                if (rId) {
                  const resolved = packageReader.resolveTarget(slidePath, rId);
                  if (resolved.path) {
                    const chartPart = packageReader.getPart(resolved.path);
                    if (chartPart && chartPart.textData) {
                      const chartParsed = ChartParser.parse(chartPart.textData, theme.colorScheme);
                      elements.push({
                        ...chartParsed,
                        id: `el-${slidePath}-${i}`,
                        name: 'Chart',
                        bounds,
                        transform,
                        zIndex: zIndexCounter++,
                        hidden: false,
                        opacity: 1,
                        source: { partPath: slidePath, xmlPath: `spTree/graphicFrame[${i}]` },
                      });
                      continue;
                    }
                  }
                }
              }

              // Fallback Unknown Graphic Frame
              elements.push({
                id: `el-${slidePath}-${i}`,
                type: 'unknown',
                name: 'Graphic Frame',
                bounds,
                transform,
                zIndex: zIndexCounter++,
                hidden: false,
                opacity: 1,
                rawXmlType: 'graphicFrame',
                fallbackText: 'SmartArt / Graphic Object',
                source: { partPath: slidePath, xmlPath: `spTree/graphicFrame[${i}]` },
              });
            }

            // 4. Connectors (p:cxnSp)
            else if (tagName === 'cxnSp') {
              const spPr = XmlUtils.findFirstTag(child, 'spPr');
              const { bounds, transform } = ShapeParser.parseTransform(spPr ? XmlUtils.findFirstTag(spPr, 'xfrm') : null);
              const line = ShapeParser.parseLine(spPr ? XmlUtils.findFirstTag(spPr, 'ln') : null, theme.colorScheme);

              elements.push({
                id: `el-${slidePath}-${i}`,
                type: 'connector',
                name: 'Connector',
                bounds,
                transform,
                zIndex: zIndexCounter++,
                hidden: false,
                opacity: 1,
                line,
                startPoint: { xEmu: bounds.xEmu, yEmu: bounds.yEmu },
                endPoint: { xEmu: bounds.xEmu + bounds.widthEmu, yEmu: bounds.yEmu + bounds.heightEmu },
                source: { partPath: slidePath, xmlPath: `spTree/cxnSp[${i}]` },
              });
            }
          }
        }

        // Notes
        let notesText = '';
        const notesRels = packageReader.getRelationships(slidePath);
        if (notesRels) {
          for (const [_, rel] of notesRels) {
            if (rel.type.includes('notesSlide')) {
              const resolved = packageReader.resolveTarget(slidePath, rel.id);
              if (resolved.path) {
                const notesPart = packageReader.getPart(resolved.path);
                if (notesPart && notesPart.textData) {
                  const notesDoc = XmlUtils.parseXml(notesPart.textData);
                  const pNodes = notesDoc.getElementsByTagName('a:p');
                  for (let n = 0; n < pNodes.length; n++) {
                    const text = pNodes[n].textContent || '';
                    if (text.trim()) notesText += text.trim() + '\n';
                  }
                }
              }
            }
          }
        }

        slides.push({
          id: `slide-${index + 1}`,
          slideNumber: index + 1,
          name: `Slide ${index + 1}`,
          background,
          elements,
          notes: notesText ? { text: notesText, paragraphs: [] } : undefined,
        });

        addDiag('info', 'SLIDE_PARSED', `Slide ${index + 1} parsed with ${elements.length} elements.`, index + 1);
      } catch (err: any) {
        addDiag('error', 'SLIDE_PARSE_ERR', `Failed to parse slide ${slidePath}: ${err.message}`, index + 1);
      }
    }

    return {
      schemaVersion: '1.0.0',
      metadata: {
        title: 'OOXML Presentation Document',
        slideCount: slides.length,
      },
      viewport,
      theme,
      slides,
      resources: resourceStore,
      diagnostics,
    };
  }
}
