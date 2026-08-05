# PPTX HTML Renderer System

A high-fidelity client-side PowerPoint (`.pptx`) OOXML parsing and rendering engine built with React, TypeScript, and Tailwind CSS. The application parses OOXML PresentationML and DrawingML parts into a normalized scene graph Intermediate Representation (IR), and renders slides using HTML, SVG, and interactive Recharts visualizations.

---

## 📌 Project Overview & Scope

### Features
- **Client-Side OOXML ZIP Engine**: Reads and unpacks `.pptx` files directly in the browser using `JSZip` without requiring external server dependencies.
- **Resilient XML Namespace Resolution**: Automatically handles missing or legacy OOXML namespace declarations (`r:`, `p:`, `a:`, `c:`, `mc:`, `v:`, `id:`), ensuring clean fallback parsing across diverse PPTX sources.
- **DrawingML & PresentationML Parser**: Parses slide structures, theme color schemes, font hierarchies, text run formatting (bold, italic, font sizes, colors), custom shape geometries (ellipses, rounded rectangles, chrevons, stars, gradients), table grids, and embedded media assets.
- **Interactive Technical Dashboard UI**: Styled following a high-density "Technical Dashboard / Data Grid" aesthetic with crisp monospaced telemetry headers, status indicators, and subtle tactile contrast.
- **Normalized IR Tree Inspector**: Live side-by-side inspection of slide, element, and theme Intermediate Representation (IR) JSON data.
- **Quality Gate Diagnostics**: Real-time validation logs and visual regression metric checks for presentation decks.
- **Presenter Studio**: Full-screen presenter mode featuring active slide stage, next-slide preview, speaker notes panel, and execution timer.
- **Standalone HTML Export**: One-click generation of self-contained HTML presentation packages with embedded SVG vectors and slide controls.

### Scope of Element Support
| Element Type | Support Level | Implementation Notes |
| :--- | :--- | :--- |
| **Text Elements** | Full | Text runs, paragraph alignments, font sizes, colors, margins, bold/italic styles |
| **Shapes & Geometries** | Full | Presets (Rect, RoundRect, Ellipse, Triangle, Chevron, Star), custom SVG paths, solid/gradient fills |
| **Tables** | Full | Column widths, row heights, cell merging (`rowSpan`/`colSpan`), custom padding, grid borders |
| **Charts** | High | Bar/Column charts, Line charts, Pie charts rendered dynamically with `Recharts` |
| **Images & Media** | Full | Base64 / Blob image rendering, HTML5 video/audio elements |
| **OLE & Embedded Data** | Partial / Extractable | OLE binary extraction and fallback UI cards |
| **Connectors** | Full | Directional line paths, dashed/dotted stroke properties |

---

## 🛠️ Build & Setup Instructions

### Prerequisites
- **Node.js**: v18.0.0 or higher
- **npm**: v9.0.0 or higher

### Installation & Development
```bash
# Install dependencies
npm install

# Start the development server (runs on port 3000)
npm run dev
```

### Build & Verification
```bash
# Type check and compile production bundle into /dist
npm run build

# Run TypeScript linter check
npm run lint

# Preview production build locally
npm run preview
```

---

## 🧪 Testing & Verification

- **Automatic Quality Gate**: Built-in `Quality Gate` panel in the application header reports total slides, total objects, parser latency metrics, and diagnostic warning/error streams.
- **Linter Check**: Execute `npm run lint` (`tsc --noEmit`) to verify strict type safety across all parser utilities and React components.
- **Build Verification**: Run `npm run build` to ensure Vite bundle optimization succeeds without missing exports or unresolved imports.

---

## 📄 Dependencies & License Summary

| Package Name | Purpose | License |
| :--- | :--- | :--- |
| `react` / `react-dom` | UI Component Framework | MIT |
| `vite` | Application Bundler & Dev Server | MIT |
| `tailwindcss` | Utility-First Styling Framework | MIT |
| `jszip` | In-Browser OOXML ZIP File Unpacking | MIT / Dual |
| `recharts` | Dynamic Slide Chart Visualizations | MIT |
| `lucide-react` | Icon Set Component Library | ISC |
| `typescript` | Static Type Checker | Apache-2.0 |
| `eslint` | Code Quality & Syntax Verification | MIT |
