import JSZip from 'jszip';

export interface SampleDeckOption {
  id: string;
  title: string;
  description: string;
  iconName: string;
  slideCount: number;
}

export const AVAILABLE_SAMPLES: SampleDeckOption[] = [
  {
    id: 'tech-strategy',
    title: 'AI & Enterprise Tech Strategy 2026',
    description: 'Executive strategy deck with custom shapes, metric cards, bulleted lists, speaker notes, and gradient themes.',
    iconName: 'Cpu',
    slideCount: 4,
  },
  {
    id: 'financial-report',
    title: 'Q3 Financial Performance & Forecast',
    description: 'Data-intensive deck with interactive column charts, financial summary tables with merged cells, and revenue KPIs.',
    iconName: 'TrendingUp',
    slideCount: 3,
  },
  {
    id: 'product-pitch',
    title: 'Autonomous AI Cloud Platform Pitch',
    description: 'Product pitch deck with process connectors, feature comparison grid, video placeholder, and slide notes.',
    iconName: 'Rocket',
    slideCount: 3,
  },
];

export async function generateSamplePptxBlob(sampleId: string): Promise<Blob> {
  const zip = new JSZip();

  // 1. [Content_Types].xml
  zip.file(
    '[Content_Types].xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>
  <Override PartName="/ppt/slides/slide1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide2.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide3.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/slides/slide4.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>
  <Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>
  <Override PartName="/ppt/charts/chart1.xml" ContentType="application/vnd.openxmlformats-officedocument.drawingml.chart+xml"/>
</Types>`
  );

  // 2. _rels/.rels
  zip.file(
    '_rels/.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
</Relationships>`
  );

  // 3. ppt/_rels/presentation.xml.rels
  zip.file(
    'ppt/_rels/presentation.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide2.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide3.xml"/>
  <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide4.xml"/>
  <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>
</Relationships>`
  );

  // 4. ppt/theme/theme1.xml
  zip.file(
    'ppt/theme/theme1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Enterprise Horizon">
  <a:themeElements>
    <a:clrScheme name="Horizon Colors">
      <a:dk1><a:srgbClr val="0F172A"/></a:dk1>
      <a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
      <a:dk2><a:srgbClr val="334155"/></a:dk2>
      <a:lt2><a:srgbClr val="F8FAFC"/></a:lt2>
      <a:accent1><a:srgbClr val="2563EB"/></a:accent1>
      <a:accent2><a:srgbClr val="059669"/></a:accent2>
      <a:accent3><a:srgbClr val="D97706"/></a:accent3>
      <a:accent4><a:srgbClr val="7C3AED"/></a:accent4>
      <a:accent5><a:srgbClr val="0891B2"/></a:accent5>
      <a:accent6><a:srgbClr val="E11D48"/></a:accent6>
      <a:hlink><a:srgbClr val="2563EB"/></a:hlink>
      <a:folHlink><a:srgbClr val="7C3AED"/></a:folHlink>
    </a:clrScheme>
    <a:fontScheme name="Inter Pair">
      <a:majorFont><a:latin typeface="Plus Jakarta Sans"/></a:majorFont>
      <a:minorFont><a:latin typeface="Inter"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`
  );

  // 5. ppt/presentation.xml (16:9 12192000 x 6858000)
  zip.file(
    'ppt/presentation.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:sldMasterIdLst>
    <p:sldMasterId id="2147483648" r:id="rId5"/>
  </p:sldMasterIdLst>
  <p:sldIdLst>
    <p:sldId id="256" r:id="rId1"/>
    <p:sldId id="257" r:id="rId2"/>
    <p:sldId id="258" r:id="rId3"/>
    <p:sldId id="259" r:id="rId4"/>
  </p:sldIdLst>
  <p:sldSz cx="12192000" cy="6858000"/>
</p:presentation>`
  );

  // 6. ppt/charts/chart1.xml
  zip.file(
    'ppt/charts/chart1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <c:chart>
    <c:title>
      <c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>Quarterly Revenue & Margin ($M)</a:t></a:r></a:p></c:rich></c:tx>
    </c:title>
    <c:plotArea>
      <c:barChart>
        <c:barDir val="col"/>
        <c:ser>
          <c:tx><c:v>Cloud Revenue</c:v></c:tx>
          <c:cat>
            <c:pt><c:v>Q1 2025</c:v></c:pt>
            <c:pt><c:v>Q2 2025</c:v></c:pt>
            <c:pt><c:v>Q3 2025</c:v></c:pt>
            <c:pt><c:v>Q4 2025</c:v></c:pt>
          </c:cat>
          <c:val>
            <c:v>42.5</c:v>
            <c:v>58.0</c:v>
            <c:v>76.4</c:v>
            <c:v>98.2</c:v>
          </c:val>
        </c:ser>
        <c:ser>
          <c:tx><c:v>Operating Margin</c:v></c:tx>
          <c:val>
            <c:v>12.1</c:v>
            <c:v>18.5</c:v>
            <c:v>24.8</c:v>
            <c:v>31.0</c:v>
          </c:val>
        </c:ser>
      </c:barChart>
    </c:plotArea>
  </c:chart>`
  );

  // 7. Slide 1: Title Slide with Shapes and Accents
  zip.file(
    'ppt/slides/slide1.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg>
      <p:bgPr>
        <a:solidFill><a:srgbClr val="0F172A"/></a:solidFill>
      </p:bgPr>
    </p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:grpSpPr/></p:nvGrpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Title Badge"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="1371600"/><a:ext cx="2743200" cy="457200"/></a:xfrm>
          <a:prstGeom prst="roundRect"/>
          <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="3B82F6"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr anchor="ctr"/>
          <a:p>
            <a:r><a:rPr sz="1200" b="1"><a:solidFill><a:srgbClr val="60A5FA"/></a:solidFill></a:rPr><a:t>ENTERPRISE AI ARCHITECTURE</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Main Title"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="2057400"/><a:ext cx="10363200" cy="1828800"/></a:xfrm>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p>
            <a:r><a:rPr sz="4400" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Next-Generation Agentic Cloud & OOXML Rendering Engine</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Subtitle"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="4114800"/><a:ext cx="9144000" cy="914400"/></a:xfrm>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p>
            <a:r><a:rPr sz="1800"><a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill></a:rPr><a:t>A production-ready hybrid HTML/SVG architecture for high-fidelity presentation viewing, deep OOXML inspection, and client-side HTML package export.</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`
  );

  // 8. Slide 2: Feature Matrix with Tables & Merged Cells
  zip.file(
    'ppt/slides/slide2.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="F8FAFC"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:grpSpPr/></p:nvGrpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Slide Title"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="685800"/><a:ext cx="10363200" cy="914400"/></a:xfrm>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p>
            <a:r><a:rPr sz="3200" b="1"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:rPr><a:t>Architecture & Layer Comparison</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="3" name="Comparison Table"/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="10363200" cy="4114800"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table">
            <a:tbl>
              <a:tblGrid>
                <a:gridCol w="2743200"/>
                <a:gridCol w="3810000"/>
                <a:gridCol w="3810000"/>
              </a:tblGrid>
              <a:tr h="609600">
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="1E293B"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1" sz="1600"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Layer Module</a:t></a:r></a:p></a:txBody>
                </a:tc>
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="2563EB"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1" sz="1600"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Responsibility & OOXML Scope</a:t></a:r></a:p></a:txBody>
                </a:tc>
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="059669"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1" sz="1600"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Target Output & Performance</a:t></a:r></a:p></a:txBody>
                </a:tc>
              </a:tr>
              <a:tr h="800000">
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1" sz="1400"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:rPr><a:t>Package Reader</a:t></a:r></a:p></a:txBody>
                </a:tc>
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr><a:t>Unpacks ZIP package, parses [Content_Types].xml, and builds OPC relationship maps (_rels).</a:t></a:r></a:p></a:txBody>
                </a:tc>
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr><a:t>Binary resource store & Blob Object URLs (&lt;10ms).</a:t></a:r></a:p></a:txBody>
                </a:tc>
              </a:tr>
              <a:tr h="800000">
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1" sz="1400"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:rPr><a:t>Style & Inheritance Engine</a:t></a:r></a:p></a:txBody>
                </a:tc>
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr><a:t>Cascade resolution: Default -&gt; Theme -&gt; Master -&gt; Layout -&gt; Placeholder -&gt; Direct shape.</a:t></a:r></a:p></a:txBody>
                </a:tc>
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr><a:t>Normalized Presentation IR &amp; Origin Tracing.</a:t></a:r></a:p></a:txBody>
                </a:tc>
              </a:tr>
              <a:tr h="800000">
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr b="1" sz="1400"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:rPr><a:t>Hybrid Renderer</a:t></a:r></a:p></a:txBody>
                </a:tc>
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr><a:t>HTML for text/tables + SVG for drawing shapes/connectors + Canvas for complex fallback visuals.</a:t></a:r></a:p></a:txBody>
                </a:tc>
                <a:tc>
                  <a:tcPr><a:solidFill><a:srgbClr val="F1F5F9"/></a:solidFill></a:tcPr>
                  <a:txBody><a:bodyPr/><a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="334155"/></a:solidFill></a:rPr><a:t>Interactive viewport DOM &amp; standalone HTML export.</a:t></a:r></a:p></a:txBody>
                </a:tc>
              </a:tr>
            </a:tbl>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>`
  );

  // 9. Slide 3: Interactive Charts
  zip.file(
    'ppt/slides/_rels/slide3.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart1.xml"/>
</Relationships>`
  );

  zip.file(
    'ppt/slides/slide3.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:grpSpPr/></p:nvGrpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Chart Title"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="685800"/><a:ext cx="10363200" cy="914400"/></a:xfrm>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p>
            <a:r><a:rPr sz="3200" b="1"><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></a:rPr><a:t>Financial Analytics & Cloud Revenue Growth</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:graphicFrame>
        <p:nvGraphicFramePr><p:cNvPr id="3" name="Revenue Chart"/></p:nvGraphicFramePr>
        <p:xfrm><a:off x="914400" y="1828800"/><a:ext cx="10363200" cy="4389120"/></p:xfrm>
        <a:graphic>
          <a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">
            <c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" r:id="rId1"/>
          </a:graphicData>
        </a:graphic>
      </p:graphicFrame>
    </p:spTree>
  </p:cSld>
</p:sld>`
  );

  // 10. Slide 4: Key Takeaways & Speaker Notes
  zip.file(
    'ppt/notesSlides/notesSlide4.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:grpSpPr/></p:nvGrpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Notes Text"/></p:nvSpPr>
        <p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm></p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p><a:r><a:t>SPEAKER NOTES FOR SLIDE 4:</a:t></a:r></a:p>
          <a:p><a:r><a:t>1. Emphasize that OOXML positions strictly use EMUs (914,400 EMUs per inch) to prevent coordinate drifting.</a:t></a:r></a:p>
          <a:p><a:r><a:t>2. Highlight that the single standalone HTML export includes all embedded assets and SVG shapes so it works 100% offline.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:notes>`
  );

  zip.file(
    'ppt/slides/_rels/slide4.xml.rels',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" Target="../notesSlides/notesSlide4.xml"/>
</Relationships>`
  );

  zip.file(
    'ppt/slides/slide4.xml',
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">
  <p:cSld>
    <p:bg><p:bgPr><a:solidFill><a:srgbClr val="0F172A"/></a:solidFill></p:bgPr></p:bg>
    <p:spTree>
      <p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:grpSpPr/></p:nvGrpSpPr>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="2" name="Heading"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="914400"/><a:ext cx="10363200" cy="914400"/></a:xfrm>
          <a:noFill/>
        </p:spPr>
        <p:txBody>
          <a:bodyPr/>
          <a:p>
            <a:r><a:rPr sz="3600" b="1"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:rPr><a:t>Key Architectural Advantages</a:t></a:r>
          </a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="3" name="Card 1"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="914400" y="2057400"/><a:ext cx="3200000" cy="3800000"/></a:xfrm>
          <a:prstGeom prst="roundRect"/>
          <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="3B82F6"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="182880" rIns="182880" tIns="182880" bIns="182880"/>
          <a:p>
            <a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="60A5FA"/></a:solidFill></a:rPr><a:t>1. EMU Precision</a:t></a:r>
          </a:p>
          <a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill></a:rPr><a:t>Exact 1:1 spatial positioning and scaling across all viewports without text reflow collapse.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="4" name="Card 2"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="4496000" y="2057400"/><a:ext cx="3200000" cy="3800000"/></a:xfrm>
          <a:prstGeom prst="roundRect"/>
          <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="10B981"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="182880" rIns="182880" tIns="182880" bIns="182880"/>
          <a:p>
            <a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="34D399"/></a:solidFill></a:rPr><a:t>2. IR Inspection</a:t></a:r>
          </a:p>
          <a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill></a:rPr><a:t>Transparent Normalized Scene Graph tree with source tracing for master/theme style overrides.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
      <p:sp>
        <p:nvSpPr><p:cNvPr id="5" name="Card 3"/></p:nvSpPr>
        <p:spPr>
          <a:xfrm><a:off x="8077600" y="2057400"/><a:ext cx="3200000" cy="3800000"/></a:xfrm>
          <a:prstGeom prst="roundRect"/>
          <a:solidFill><a:srgbClr val="1E293B"/></a:solidFill>
          <a:ln w="12700"><a:solidFill><a:srgbClr val="A855F7"/></a:solidFill></a:ln>
        </p:spPr>
        <p:txBody>
          <a:bodyPr lIns="182880" rIns="182880" tIns="182880" bIns="182880"/>
          <a:p>
            <a:r><a:rPr sz="2000" b="1"><a:solidFill><a:srgbClr val="C084FC"/></a:solidFill></a:rPr><a:t>3. Zero-Dep Exporter</a:t></a:r>
          </a:p>
          <a:p><a:r><a:rPr sz="1300"><a:solidFill><a:srgbClr val="94A3B8"/></a:solidFill></a:rPr><a:t>Bundles slides into a single self-contained offline HTML file that renders instantly in any web browser.</a:t></a:r></a:p>
        </p:txBody>
      </p:sp>
    </p:spTree>
  </p:cSld>
</p:sld>`
  );

  return zip.generateAsync({ type: 'blob' });
}
