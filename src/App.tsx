import React, { useEffect, useState } from 'react';
import { PresentationDocument, SlideElement } from './types/pptx';
import { OoxmlPackageReader } from './engine/package/packageReader';
import { SlideParser } from './engine/parser/slideParser';
import { generateSamplePptxBlob } from './engine/sample/sampleGenerator';
import { Header } from './components/Header';
import { ThumbnailsBar } from './components/ThumbnailsBar';
import { SlideViewport } from './components/SlideViewport';
import { IrInspector } from './components/IrInspector';
import { DiagnosticsPanel } from './components/DiagnosticsPanel';
import { PresenterMode } from './components/PresenterMode';
import { ExportModal } from './components/ExportModal';
import { Loader2, UploadCloud, ChevronLeft, ChevronRight, Maximize2 } from 'lucide-react';

export default function App() {
  const [packageReader] = useState(() => new OoxmlPackageReader());
  const [presentation, setPresentation] = useState<PresentationDocument | null>(null);
  const [currentSlideIndex, setCurrentSlideIndex] = useState(0);
  const [selectedElement, setSelectedElement] = useState<SlideElement | null>(null);
  const [scaleMode, setScaleMode] = useState<'fit' | '100' | '125' | '150'>('fit');
  const [searchQuery, setSearchQuery] = useState('');

  // UI Drawer / Modal Toggles
  const [inspectorOpen, setInspectorOpen] = useState(false);
  const [diagnosticsOpen, setDiagnosticsOpen] = useState(false);
  const [presenterOpen, setPresenterOpen] = useState(false);
  const [exportOpen, setExportOpen] = useState(false);

  // Loading States
  const [isLoading, setIsLoading] = useState(true);
  const [loadingText, setLoadingText] = useState('Initializing OOXML Engine...');
  const [isDragging, setIsDragging] = useState(false);

  // Load a PPTX File (Blob or File)
  const loadPptxFile = async (fileOrBlob: File | Blob, titleHint?: string) => {
    setIsLoading(true);
    setLoadingText('Unpacking OOXML ZIP Package...');
    setSelectedElement(null);

    try {
      const resourceStore = await packageReader.load(fileOrBlob);
      setLoadingText('Parsing PresentationML & DrawingML IR...');
      const parsedPres = await SlideParser.parsePresentation(packageReader, resourceStore);

      if (titleHint) {
        parsedPres.metadata.title = titleHint;
      } else if (fileOrBlob instanceof File) {
        parsedPres.metadata.title = fileOrBlob.name.replace(/\.pptx$/i, '');
      }

      setPresentation(parsedPres);
      setCurrentSlideIndex(0);
    } catch (err: any) {
      console.error('Failed to parse PPTX:', err);
    } finally {
      setIsLoading(false);
    }
  };

  // Load Initial Default Sample Deck
  useEffect(() => {
    const initDefaultSample = async () => {
      const blob = await generateSamplePptxBlob('tech-strategy');
      await loadPptxFile(blob, 'AI & Tech Strategy 2026');
    };
    initDefaultSample();

    return () => {
      packageReader.dispose();
    };
  }, []);

  // Keyboard navigation
  useEffect(() => {
    const handleKeyDown = (e: KeyboardEvent) => {
      if (presenterOpen || !presentation) return;
      if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement) return;

      if (e.key === 'ArrowRight' || e.key === 'Space') {
        if (currentSlideIndex < presentation.slides.length - 1) {
          setCurrentSlideIndex((prev) => prev + 1);
        }
      } else if (e.key === 'ArrowLeft') {
        if (currentSlideIndex > 0) {
          setCurrentSlideIndex((prev) => prev - 1);
        }
      }
    };

    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [presentation, currentSlideIndex, presenterOpen]);

  // Drag and drop file upload handlers
  const handleDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(true);
  };

  const handleDragLeave = () => {
    setIsDragging(false);
  };

  const handleDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file && file.name.endsWith('.pptx')) {
      await loadPptxFile(file);
    }
  };

  const currentSlide = presentation?.slides[currentSlideIndex];

  return (
    <div
      className="flex flex-col h-screen w-screen bg-[#E4E3E0] text-[#141414] overflow-hidden font-mono select-none"
      onDragOver={handleDragOver}
      onDragLeave={handleDragLeave}
      onDrop={handleDrop}
    >
      {/* Header */}
      <Header
        title={presentation?.metadata.title || ''}
        onLoadSample={async (id) => {
          const blob = await generateSamplePptxBlob(id);
          const sample = {
            'tech-strategy': 'AI & Tech Strategy 2026',
            'financial-report': 'Q3 Financial Performance Report',
            'product-pitch': 'Autonomous AI Cloud Pitch',
          }[id] || 'Sample Presentation';
          await loadPptxFile(blob, sample);
        }}
        onFileUpload={loadPptxFile}
        onTogglePresenter={() => setPresenterOpen(true)}
        onToggleDiagnostics={() => setDiagnosticsOpen(true)}
        onToggleInspector={() => setInspectorOpen(!inspectorOpen)}
        onOpenExport={() => setExportOpen(true)}
        inspectorOpen={inspectorOpen}
        isLoading={isLoading}
      />

      {/* Main Workspace Body */}
      <div className="flex-1 flex overflow-hidden relative">
        {isLoading && (
          <div className="absolute inset-0 z-50 bg-[#E4E3E0]/95 backdrop-blur-xs flex flex-col items-center justify-center gap-3">
            <Loader2 className="w-8 h-8 text-[#141414] animate-spin" />
            <p className="text-xs font-mono font-bold uppercase tracking-wider text-[#141414]">// {loadingText}</p>
          </div>
        )}

        {/* Drag and Drop File Overlay */}
        {isDragging && (
          <div className="absolute inset-0 z-50 bg-[#141414]/80 backdrop-blur border-4 border-dashed border-[#E4E3E0] flex flex-col items-center justify-center gap-2 pointer-events-none">
            <UploadCloud className="w-12 h-12 text-[#E4E3E0] animate-bounce" />
            <p className="text-sm font-mono font-bold text-[#E4E3E0] uppercase tracking-wider">
              [DROP .PPTX DECK TO PARSE]
            </p>
          </div>
        )}

        {presentation && (
          <>
            {/* Left Thumbnails Navigation Bar */}
            <ThumbnailsBar
              presentation={presentation}
              currentSlideIndex={currentSlideIndex}
              onSelectSlide={(idx) => {
                setCurrentSlideIndex(idx);
                setSelectedElement(null);
              }}
              searchQuery={searchQuery}
              onSearchChange={setSearchQuery}
            />

            {/* Center Stage Viewport */}
            <main className="flex-1 flex flex-col relative overflow-hidden bg-[#D9D7D1]">
              {currentSlide && (
                <SlideViewport
                  slide={currentSlide}
                  viewport={presentation.viewport}
                  selectedElement={selectedElement}
                  onSelectElement={setSelectedElement}
                  scaleMode={scaleMode}
                  onScaleChange={setScaleMode}
                />
              )}

              {/* Bottom Floating Navigation Controls */}
              <div className="absolute bottom-4 left-1/2 -translate-x-1/2 z-30 bg-[#E4E3E0] border border-[#141414] px-4 py-1.5 flex items-center gap-4 shadow-xl">
                <button
                  disabled={currentSlideIndex === 0}
                  onClick={() => {
                    setCurrentSlideIndex((prev) => Math.max(0, prev - 1));
                    setSelectedElement(null);
                  }}
                  className="p-1 hover:bg-[#141414] hover:text-[#E4E3E0] border border-[#141414] text-[#141414] disabled:opacity-30 transition-colors"
                  title="Previous Slide (←)"
                >
                  <ChevronLeft className="w-4 h-4" />
                </button>

                <span className="text-xs font-mono font-bold text-[#141414]">
                  SLD_{String(currentSlideIndex + 1).padStart(2, '0')} / SLD_
                  {String(presentation.slides.length).padStart(2, '0')}
                </span>

                <button
                  disabled={currentSlideIndex === presentation.slides.length - 1}
                  onClick={() => {
                    setCurrentSlideIndex((prev) => Math.min(presentation.slides.length - 1, prev + 1));
                    setSelectedElement(null);
                  }}
                  className="p-1 hover:bg-[#141414] hover:text-[#E4E3E0] border border-[#141414] text-[#141414] disabled:opacity-30 transition-colors"
                  title="Next Slide (→)"
                >
                  <ChevronRight className="w-4 h-4" />
                </button>
              </div>
            </main>

            {/* Right Normalized IR Inspector Drawer */}
            {inspectorOpen && (
              <IrInspector
                presentation={presentation}
                selectedElement={selectedElement}
                currentSlideIndex={currentSlideIndex}
              />
            )}
          </>
        )}
      </div>

      {/* Presenter Mode Modal Overlay */}
      {presenterOpen && presentation && (
        <PresenterMode
          presentation={presentation}
          currentSlideIndex={currentSlideIndex}
          onSelectSlide={setCurrentSlideIndex}
          onClose={() => setPresenterOpen(false)}
        />
      )}

      {/* Diagnostics & Quality Gate Modal */}
      {diagnosticsOpen && presentation && (
        <DiagnosticsPanel
          presentation={presentation}
          onClose={() => setDiagnosticsOpen(false)}
        />
      )}

      {/* Standalone HTML Export Modal */}
      {exportOpen && presentation && (
        <ExportModal
          presentation={presentation}
          onClose={() => setExportOpen(false)}
        />
      )}
    </div>
  );
}
