import { PresentationDocument, SlideDocument, SlideElement } from '../../types/pptx';
import { emuToPx } from '../utils/emu';

export class HtmlExporter {
  static exportStandaloneHtml(presentation: PresentationDocument): string {
    const slideWidthPx = Math.round(emuToPx(presentation.viewport.widthEmu));
    const slideHeightPx = Math.round(emuToPx(presentation.viewport.heightEmu));

    const slidesHtml = presentation.slides
      .map((slide, idx) => HtmlExporter.renderSlideHtml(slide, slideWidthPx, slideHeightPx, idx === 0))
      .join('\n');

    return `<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${presentation.metadata.title || 'Exported Presentation'}</title>
  <style>
    * { box-sizing: border-box; margin: 0; padding: 0; }
    body {
      background-color: #0F172A;
      color: #F8FAFC;
      font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, Helvetica, Arial, sans-serif;
      display: flex;
      flex-direction: column;
      height: 100vh;
      overflow: hidden;
    }
    .deck-header {
      height: 48px;
      background-color: #1E293B;
      border-bottom: 1px solid #334155;
      display: flex;
      align-items: center;
      justify-content: space-between;
      padding: 0 16px;
      font-size: 14px;
    }
    .deck-controls {
      display: flex;
      gap: 8px;
    }
    .deck-btn {
      background-color: #3B82F6;
      color: white;
      border: none;
      padding: 6px 12px;
      border-radius: 4px;
      cursor: pointer;
      font-weight: 500;
    }
    .deck-btn:hover { background-color: #2563EB; }
    .deck-body {
      flex: 1;
      display: flex;
      align-items: center;
      justify-content: center;
      position: relative;
      overflow: hidden;
      padding: 20px;
    }
    .slide-wrapper {
      position: absolute;
      width: ${slideWidthPx}px;
      height: ${slideHeightPx}px;
      box-shadow: 0 25px 50px -12px rgba(0, 0, 0, 0.5);
      border-radius: 8px;
      overflow: hidden;
      display: none;
      transform-origin: center center;
    }
    .slide-wrapper.active {
      display: block;
    }
    .pptx-text {
      position: absolute;
      overflow: hidden;
      word-break: break-word;
    }
    .pptx-shape {
      position: absolute;
    }
    .pptx-table {
      position: absolute;
      border-collapse: collapse;
    }
    .pptx-table td {
      border: 1px solid #CBD5E1;
      padding: 6px;
    }
  </style>
</head>
<body>
  <div class="deck-header">
    <div><strong>${presentation.metadata.title || 'Presentation Deck'}</strong> (${presentation.slides.length} Slides)</div>
    <div class="deck-controls">
      <button class="deck-btn" onclick="prevSlide()">← Prev</button>
      <span id="slide-num">Slide 1 / ${presentation.slides.length}</span>
      <button class="deck-btn" onclick="nextSlide()">Next →</button>
    </div>
  </div>
  <div class="deck-body" id="stage">
    ${slidesHtml}
  </div>

  <script>
    let currentSlide = 0;
    const totalSlides = ${presentation.slides.length};
    const slides = document.querySelectorAll('.slide-wrapper');
    const numDisplay = document.getElementById('slide-num');

    function showSlide(idx) {
      slides.forEach((s, i) => {
        if (i === idx) s.classList.add('active');
        else s.classList.remove('active');
      });
      currentSlide = idx;
      numDisplay.textContent = \`Slide \${idx + 1} / \${totalSlides}\`;
      fitScale();
    }

    function prevSlide() {
      if (currentSlide > 0) showSlide(currentSlide - 1);
    }

    function nextSlide() {
      if (currentSlide < totalSlides - 1) showSlide(currentSlide + 1);
    }

    function fitScale() {
      const stage = document.getElementById('stage');
      const w = ${slideWidthPx};
      const h = ${slideHeightPx};
      const scale = Math.min((stage.clientWidth - 40) / w, (stage.clientHeight - 40) / h);
      slides.forEach(s => {
        s.style.transform = \`scale(\${scale})\`;
      });
    }

    document.addEventListener('keydown', (e) => {
      if (e.key === 'ArrowRight' || e.key === 'Space') nextSlide();
      if (e.key === 'ArrowLeft') prevSlide();
    });

    window.addEventListener('resize', fitScale);
    fitScale();
  </script>
</body>
</html>`;
  }

  private static renderSlideHtml(
    slide: SlideDocument,
    widthPx: number,
    heightPx: number,
    isActive: boolean
  ): string {
    const bgStyle =
      slide.background.type === 'solid'
        ? `background-color: ${slide.background.color};`
        : slide.background.type === 'gradient'
        ? `background: linear-gradient(${slide.background.angleDegrees}deg, ${slide.background.stops.map((s) => `${s.color} ${Math.round(s.position * 100)}%`).join(', ')});`
        : `background-color: #FFFFFF;`;

    const elementsHtml = slide.elements.map((el) => HtmlExporter.renderElementHtml(el)).join('\n');

    return `<div class="slide-wrapper ${isActive ? 'active' : ''}" style="width: ${widthPx}px; height: ${heightPx}px; ${bgStyle}">
      ${elementsHtml}
    </div>`;
  }

  private static renderElementHtml(el: SlideElement): string {
    const xPx = emuToPx(el.bounds.xEmu);
    const yPx = emuToPx(el.bounds.yEmu);
    const wPx = emuToPx(el.bounds.widthEmu);
    const hPx = emuToPx(el.bounds.heightEmu);

    const posStyle = `left: ${xPx}px; top: ${yPx}px; width: ${wPx}px; height: ${hPx}px; transform: rotate(${el.transform.rotationDegrees}deg); z-index: ${el.zIndex};`;

    if (el.type === 'text') {
      const textContent = el.paragraphs
        .map((p) => {
          const runs = p.runs
            .map((r) => {
              const fontStyle = `font-family: ${r.style.fontFamily}; font-size: ${r.style.fontSizePt}pt; color: ${r.style.color}; ${r.style.bold ? 'font-weight: bold;' : ''} ${r.style.italic ? 'font-style: italic;' : ''}`;
              return `<span style="${fontStyle}">${r.text}</span>`;
            })
            .join('');
          return `<div style="text-align: ${p.alignment}; margin-bottom: 4px;">${runs}</div>`;
        })
        .join('');

      return `<div class="pptx-text" style="${posStyle} padding: ${el.bodyStyle.paddingPx.top}px ${el.bodyStyle.paddingPx.right}px ${el.bodyStyle.paddingPx.bottom}px ${el.bodyStyle.paddingPx.left}px;">${textContent}</div>`;
    }

    if (el.type === 'shape') {
      const fillStyle =
        el.fill.type === 'solid'
          ? `background-color: ${el.fill.color};`
          : el.fill.type === 'gradient'
          ? `background: linear-gradient(${el.fill.angleDegrees}deg, ${el.fill.stops.map((s) => `${s.color} ${Math.round(s.position * 100)}%`).join(', ')});`
          : 'background-color: transparent;';

      const borderStyle = el.line.widthPx > 0 ? `border: ${el.line.widthPx}px ${el.line.dash === 'dashed' ? 'dashed' : 'solid'} ${el.line.color};` : '';
      const borderRadius = el.geometry.presetName === 'roundRect' ? 'border-radius: 12px;' : el.geometry.presetName === 'ellipse' ? 'border-radius: 50%;' : '';

      return `<div class="pptx-shape" style="${posStyle} ${fillStyle} ${borderStyle} ${borderRadius}"></div>`;
    }

    if (el.type === 'table') {
      const rowsHtml = el.rows
        .map((row) => {
          const cells = row.cells
            .map((cell) => {
              const bg = cell.fill.type === 'solid' ? `background-color: ${cell.fill.color};` : '';
              return `<td style="${bg}">${cell.text.paragraphs.map((p) => p.runs.map((r) => r.text).join('')).join('')}</td>`;
            })
            .join('');
          return `<tr>${cells}</tr>`;
        })
        .join('');

      return `<table class="pptx-table" style="${posStyle}">${rowsHtml}</table>`;
    }

    return `<div style="${posStyle} background: rgba(255,255,255,0.1); border: 1px dashed #64748B;"></div>`;
  }
}
