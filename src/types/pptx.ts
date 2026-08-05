/**
 * OOXML PPTX Normalized Presentation Intermediate Representation (IR)
 * Independent from OOXML XML structure and browser DOM.
 */

export interface BoundingBox {
  xEmu: number;
  yEmu: number;
  widthEmu: number;
  heightEmu: number;
}

export interface Transform2D {
  rotationDegrees: number;
  flipHorizontal: boolean;
  flipVertical: boolean;
  scaleX?: number;
  scaleY?: number;
}

export type ElementType =
  | 'text'
  | 'shape'
  | 'image'
  | 'table'
  | 'chart'
  | 'group'
  | 'media'
  | 'ole'
  | 'connector'
  | 'unknown';

export interface SourceReference {
  partPath: string;
  xmlPath?: string;
  elementId?: string;
}

export interface BaseElement {
  id: string;
  type: ElementType;
  name?: string;
  bounds: BoundingBox;
  transform: Transform2D;
  zIndex: number;
  hidden: boolean;
  opacity: number;
  source: SourceReference;
  hyperlinkUrl?: string;
}

export type FillType = 'none' | 'solid' | 'gradient' | 'picture' | 'pattern';

export interface SolidFill {
  type: 'solid';
  color: string; // hex or rgba
  opacity: number;
}

export interface GradientStop {
  position: number; // 0 to 1
  color: string;
}

export interface GradientFill {
  type: 'gradient';
  angleDegrees: number;
  stops: GradientStop[];
}

export interface PictureFill {
  type: 'picture';
  resourceId: string;
  mimeType: string;
  srcUrl: string;
}

export interface NoFill {
  type: 'none';
}

export type FillStyle = SolidFill | GradientFill | PictureFill | NoFill;

export type LineDash = 'solid' | 'dashed' | 'dotted' | 'none';

export interface LineStyle {
  color: string;
  widthPx: number;
  dash: LineDash;
  cap?: 'flat' | 'square' | 'round';
  startArrow?: boolean;
  endArrow?: boolean;
}

// Text Types
export type VerticalAlignment = 'top' | 'middle' | 'bottom';

export interface TextBodyStyle {
  paddingPx: { top: number; right: number; bottom: number; left: number };
  verticalAlign: VerticalAlignment;
  wrapText: boolean;
  autofit: 'none' | 'shrink' | 'resizeShape';
}

export type TextAlignment = 'left' | 'center' | 'right' | 'justify';

export interface BulletStyle {
  type: 'none' | 'char' | 'number';
  char?: string;
  startAt?: number;
  color?: string;
}

export interface TextRunStyle {
  fontFamily: string;
  fontSizePt: number;
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strikethrough: boolean;
  color: string;
  backgroundColor?: string;
  superscript?: boolean;
  subscript?: boolean;
  letterSpacingPt?: number;
  hyperlinkUrl?: string;
}

export interface TextRun {
  type: 'text' | 'break' | 'field';
  text: string;
  fieldType?: string;
  style: TextRunStyle;
}

export interface TextParagraph {
  id: string;
  alignment: TextAlignment;
  spaceBeforePt?: number;
  spaceAfterPt?: number;
  lineSpacingPct?: number;
  indentLevel: number;
  bullet?: BulletStyle;
  runs: TextRun[];
}

export interface TextElement extends BaseElement {
  type: 'text';
  bodyStyle: TextBodyStyle;
  paragraphs: TextParagraph[];
  fillStyle?: FillStyle;
}

// Shape Types
export interface ShapeGeometry {
  kind: 'preset' | 'custom';
  presetName?: string; // e.g. rect, roundRect, ellipse, triangle, chevron, star5, callout1
  svgPath?: string; // custom or compiled path
  cornerRadiusPx?: number;
}

export interface ShapeElement extends BaseElement {
  type: 'shape';
  geometry: ShapeGeometry;
  fill: FillStyle;
  line: LineStyle;
  text?: {
    bodyStyle: TextBodyStyle;
    paragraphs: TextParagraph[];
  };
  shadow?: {
    color: string;
    blurPx: number;
    offsetX: number;
    offsetY: number;
  };
}

// Image Types
export interface ImageElement extends BaseElement {
  type: 'image';
  resourceId: string;
  mimeType: string;
  srcUrl: string;
  crop?: { topPct: number; rightPct: number; bottomPct: number; leftPct: number };
  altText?: string;
}

// Table Types
export interface TableCell {
  id: string;
  rowSpan: number;
  colSpan: number;
  fill: FillStyle;
  borders: {
    top?: LineStyle;
    right?: LineStyle;
    bottom?: LineStyle;
    left?: LineStyle;
  };
  paddingPx: { top: number; right: number; bottom: number; left: number };
  verticalAlign: VerticalAlignment;
  text: {
    paragraphs: TextParagraph[];
  };
  hidden?: boolean; // Continuation of rowSpan/colSpan
}

export interface TableRow {
  heightEmu: number;
  cells: TableCell[];
}

export interface TableColumn {
  widthEmu: number;
}

export interface TableElement extends BaseElement {
  type: 'table';
  columns: TableColumn[];
  rows: TableRow[];
}

// Chart Types
export type ChartType = 'bar' | 'column' | 'line' | 'pie' | 'area';

export interface ChartSeries {
  name: string;
  color: string;
  values: number[];
}

export interface ChartElement extends BaseElement {
  type: 'chart';
  title?: string;
  chartType: ChartType;
  categories: string[];
  series: ChartSeries[];
  showLegend: boolean;
  showGridLines: boolean;
}

// Group Types
export interface GroupCoordinateSpace {
  childOffsetXEmu: number;
  childOffsetYEmu: number;
  childWidthEmu: number;
  childHeightEmu: number;
}

export interface GroupElement extends BaseElement {
  type: 'group';
  coordSpace: GroupCoordinateSpace;
  children: SlideElement[];
}

// Media & OLE Types
export interface MediaElement extends BaseElement {
  type: 'media';
  mediaKind: 'audio' | 'video';
  srcUrl?: string;
  posterUrl?: string;
  autoPlay: boolean;
  loop: boolean;
}

export interface OleElement extends BaseElement {
  type: 'ole';
  programId?: string;
  displayMode: 'icon' | 'preview';
  previewUrl?: string;
  fileName?: string;
  fileData?: Uint8Array;
}

export interface ConnectorElement extends BaseElement {
  type: 'connector';
  line: LineStyle;
  startPoint: { xEmu: number; yEmu: number };
  endPoint: { xEmu: number; yEmu: number };
  presetKind?: 'straight' | 'bent' | 'curved';
}

export interface UnknownElement extends BaseElement {
  type: 'unknown';
  rawXmlType: string;
  fallbackText?: string;
}

export type SlideElement =
  | TextElement
  | ShapeElement
  | ImageElement
  | TableElement
  | ChartElement
  | GroupElement
  | MediaElement
  | OleElement
  | ConnectorElement
  | UnknownElement;

export interface SlideViewport {
  widthEmu: number;
  heightEmu: number;
  aspectRatio: number;
}

export interface ColorScheme {
  dk1: string;
  lt1: string;
  dk2: string;
  lt2: string;
  accent1: string;
  accent2: string;
  accent3: string;
  accent4: string;
  accent5: string;
  accent6: string;
  hlink: string;
  folHlink: string;
  [key: string]: string;
}

export interface FontScheme {
  majorFont: string;
  minorFont: string;
}

export interface ResolvedTheme {
  id: string;
  name: string;
  colorScheme: ColorScheme;
  fontScheme: FontScheme;
}

export interface NotesDocument {
  text: string;
  paragraphs: TextParagraph[];
}

export interface SlideDocument {
  id: string;
  slideNumber: number;
  name?: string;
  background: FillStyle;
  elements: SlideElement[];
  notes?: NotesDocument;
  masterId?: string;
  layoutId?: string;
  layoutName?: string;
}

export interface PresentationMetadata {
  title?: string;
  subject?: string;
  creator?: string;
  lastModifiedBy?: string;
  createdDate?: string;
  slideCount: number;
}

export type DiagnosticSeverity = 'info' | 'warning' | 'error';

export interface Diagnostic {
  id: string;
  severity: DiagnosticSeverity;
  code: string;
  message: string;
  slideIndex?: number;
  elementId?: string;
  timestamp: number;
}

export interface ResourceStore {
  images: Map<string, { mimeType: string; url: string; data: Uint8Array }>;
  media: Map<string, { mimeType: string; url: string }>;
  embeddedFiles: Map<string, { fileName: string; data: Uint8Array }>;
}

export interface PresentationDocument {
  schemaVersion: string;
  metadata: PresentationMetadata;
  viewport: SlideViewport;
  theme: ResolvedTheme;
  slides: SlideDocument[];
  resources: ResourceStore;
  diagnostics: Diagnostic[];
}

export type StyleOrigin = 'default' | 'theme' | 'master' | 'layout' | 'placeholder' | 'direct';

export interface StyleValueTrace<T> {
  value: T;
  origin: StyleOrigin;
}
