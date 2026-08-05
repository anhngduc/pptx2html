import React, { useEffect, useState } from 'react';
import { PresentationDocument } from '../types/pptx';
import { SlideViewport } from './SlideViewport';
import { Clock, FileText, Play, RotateCcw, X, ChevronLeft, ChevronRight } from 'lucide-react';

interface Props {
  presentation: PresentationDocument;
  currentSlideIndex: number;
  onSelectSlide: (index: number) => void;
  onClose: () => void;
}

export const PresenterMode: React.FC<Props> = ({
  presentation,
  currentSlideIndex,
  onSelectSlide,
  onClose,
}) => {
  const [seconds, setSeconds] = useState(0);
  const [timerRunning, setTimerRunning] = useState(true);

  useEffect(() => {
    let interval: any = null;
    if (timerRunning) {
      interval = setInterval(() => setSeconds((s) => s + 1), 1000);
    }
    return () => clearInterval(interval);
  }, [timerRunning]);

  const formatTimer = (totalSecs: number) => {
    const mins = Math.floor(totalSecs / 60);
    const secs = totalSecs % 60;
    return `${mins.toString().padStart(2, '0')}:${secs.toString().padStart(2, '0')}`;
  };

  const currentSlide = presentation.slides[currentSlideIndex];
  const nextSlide = presentation.slides[currentSlideIndex + 1];

  const handleKeyDown = (e: KeyboardEvent) => {
    if (e.key === 'ArrowRight' || e.key === 'Space') {
      if (currentSlideIndex < presentation.slides.length - 1) {
        onSelectSlide(currentSlideIndex + 1);
      }
    } else if (e.key === 'ArrowLeft') {
      if (currentSlideIndex > 0) {
        onSelectSlide(currentSlideIndex - 1);
      }
    } else if (e.key === 'Escape') {
      onClose();
    }
  };

  useEffect(() => {
    window.addEventListener('keydown', handleKeyDown);
    return () => window.removeEventListener('keydown', handleKeyDown);
  }, [currentSlideIndex]);

  return (
    <div className="fixed inset-0 z-50 bg-[#141414] text-[#E4E3E0] flex flex-col select-none font-mono">
      {/* Top Controls Bar */}
      <div className="h-14 bg-[#E4E3E0] border-b border-[#141414] px-6 flex items-center justify-between text-[#141414]">
        <div className="flex items-center gap-4">
          <span className="font-bold text-xs uppercase tracking-wider bg-[#141414] text-[#E4E3E0] px-2 py-0.5">
            PRESENTER_CONTROL_ROOM
          </span>
          <span className="text-xs font-serif italic text-stone-600">
            SLD_{String(currentSlideIndex + 1).padStart(2, '0')} / SLD_{String(presentation.slides.length).padStart(2, '0')}
          </span>
        </div>

        {/* Timer */}
        <div className="flex items-center gap-3 bg-[#F0EFEA] px-3 py-1 border border-[#141414] text-xs">
          <Clock className="w-3.5 h-3.5 text-[#141414]" />
          <span className="font-mono text-xs font-bold text-[#141414]">{formatTimer(seconds)}</span>
          <button
            onClick={() => setTimerRunning(!timerRunning)}
            className="text-[10px] uppercase font-bold text-[#141414] hover:underline"
          >
            [{timerRunning ? 'PAUSE' : 'RESUME'}]
          </button>
          <button
            onClick={() => setSeconds(0)}
            className="p-0.5 text-[#141414] hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors"
            title="Reset Timer"
          >
            <RotateCcw className="w-3 h-3" />
          </button>
        </div>

        {/* Slide Navigation & Exit */}
        <div className="flex items-center gap-2">
          <button
            disabled={currentSlideIndex === 0}
            onClick={() => onSelectSlide(currentSlideIndex - 1)}
            className="p-1.5 bg-[#E4E3E0] hover:bg-[#141414] hover:text-[#E4E3E0] border border-[#141414] disabled:opacity-40 transition-colors text-[#141414]"
          >
            <ChevronLeft className="w-4 h-4" />
          </button>
          <button
            disabled={currentSlideIndex === presentation.slides.length - 1}
            onClick={() => onSelectSlide(currentSlideIndex + 1)}
            className="p-1.5 bg-[#E4E3E0] hover:bg-[#141414] hover:text-[#E4E3E0] border border-[#141414] disabled:opacity-40 transition-colors text-[#141414]"
          >
            <ChevronRight className="w-4 h-4" />
          </button>
          <button
            onClick={onClose}
            className="p-1.5 bg-[#D73A49] hover:bg-red-800 text-white border border-[#141414] ml-2 transition-colors"
            title="Exit Presenter Mode (Esc)"
          >
            <X className="w-4 h-4" />
          </button>
        </div>
      </div>

      {/* Main Dual Stage View */}
      <div className="flex-1 grid grid-cols-3 gap-4 p-6 overflow-hidden bg-[#141414]">
        {/* Main Current Slide Stage */}
        <div className="col-span-2 bg-[#E4E3E0] border-2 border-[#141414] overflow-hidden flex flex-col p-4 relative text-[#141414]">
          <span className="absolute top-2 left-3 z-10 text-[10px] font-serif italic uppercase text-stone-600 tracking-wider">
            // ACTIVE_VIEWPORT
          </span>
          {currentSlide && (
            <SlideViewport
              slide={currentSlide}
              viewport={presentation.viewport}
              selectedElement={null}
              onSelectElement={() => {}}
              scaleMode="fit"
              onScaleChange={() => {}}
            />
          )}
        </div>

        {/* Sidebar: Next Slide Preview & Speaker Notes */}
        <div className="col-span-1 flex flex-col gap-4">
          {/* Next Slide Preview */}
          <div className="flex-1 bg-[#E4E3E0] border-2 border-[#141414] p-3 flex flex-col relative overflow-hidden text-[#141414]">
            <span className="text-[10px] font-serif italic uppercase text-stone-600 tracking-wider mb-2">
              // QUEUED_STAGE_PREVIEW
            </span>
            {nextSlide ? (
              <SlideViewport
                slide={nextSlide}
                viewport={presentation.viewport}
                selectedElement={null}
                onSelectElement={() => {}}
                scaleMode="fit"
                onScaleChange={() => {}}
              />
            ) : (
              <div className="flex-1 flex items-center justify-center text-stone-500 font-mono text-xs uppercase">
                // END_OF_DECK
              </div>
            )}
          </div>

          {/* Speaker Notes Drawer */}
          <div className="h-48 bg-[#E4E3E0] border-2 border-[#141414] p-4 flex flex-col text-[#141414]">
            <div className="flex items-center gap-2 mb-2 pb-1 border-b border-[#141414]">
              <FileText className="w-4 h-4 text-amber-700" />
              <span className="font-bold text-xs uppercase tracking-wide">// SPEAKER_NOTES</span>
            </div>
            <div className="flex-1 overflow-y-auto text-xs text-stone-800 font-mono leading-relaxed whitespace-pre-wrap bg-[#F0EFEA] p-2 border border-[#141414]/30">
              {currentSlide?.notes?.text || 'No speaker notes for this slide.'}
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
