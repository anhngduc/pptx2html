import React from 'react';
import { PresentationDocument, SlideDocument } from '../types/pptx';
import { FileText, Layers, Search } from 'lucide-react';

interface Props {
  presentation: PresentationDocument;
  currentSlideIndex: number;
  onSelectSlide: (index: number) => void;
  searchQuery: string;
  onSearchChange: (q: string) => void;
}

export const ThumbnailsBar: React.FC<Props> = ({
  presentation,
  currentSlideIndex,
  onSelectSlide,
  searchQuery,
  onSearchChange,
}) => {
  const filteredSlides = presentation.slides.filter((slide) => {
    if (!searchQuery) return true;
    const q = searchQuery.toLowerCase();
    return (
      slide.name?.toLowerCase().includes(q) ||
      slide.notes?.text.toLowerCase().includes(q) ||
      slide.elements.some((el) => el.name?.toLowerCase().includes(q))
    );
  });

  return (
    <div className="w-64 h-full bg-[#E4E3E0] border-r border-[#141414] flex flex-col select-none text-[#141414]">
      {/* Search Header */}
      <div className="p-3 border-b border-[#141414] bg-[#D9D7D1]">
        <div className="font-serif italic text-[11px] text-stone-600 uppercase mb-1.5 flex items-center justify-between">
          <span>// SLIDE_SLOTS</span>
          <span className="font-mono text-[10px] bg-[#141414] text-[#E4E3E0] px-1.5 py-0.2">{presentation.slides.length}</span>
        </div>
        <div className="relative">
          <Search className="w-3.5 h-3.5 absolute left-2.5 top-2 text-[#141414]" />
          <input
            type="text"
            placeholder="Search slides, notes..."
            value={searchQuery}
            onChange={(e) => onSearchChange(e.target.value)}
            className="w-full bg-[#F0EFEA] border border-[#141414] pl-8 pr-2 py-1 text-xs font-mono text-[#141414] placeholder-stone-500 focus:outline-none focus:bg-white"
          />
        </div>
      </div>

      {/* Thumbnails List */}
      <div className="flex-1 overflow-y-auto p-3 space-y-2.5">
        {filteredSlides.map((slide, idx) => {
          const originalIndex = presentation.slides.findIndex((s) => s.id === slide.id);
          const isSelected = originalIndex === currentSlideIndex;

          return (
            <div
              key={slide.id}
              onClick={() => onSelectSlide(originalIndex)}
              className={`group cursor-pointer border p-2 transition-all ${
                isSelected
                  ? 'bg-[#141414] text-[#E4E3E0] border-[#141414]'
                  : 'bg-[#F0EFEA] border-[#141414] hover:bg-white text-[#141414]'
              }`}
            >
              {/* Thumbnail Container */}
              <div
                className={`relative aspect-video border overflow-hidden flex items-center justify-center mb-1.5 ${
                  isSelected ? 'bg-stone-900 border-[#E4E3E0]/30' : 'bg-white border-[#141414]'
                }`}
              >
                <div className={`text-[10px] font-mono flex flex-col items-center ${isSelected ? 'text-stone-300' : 'text-stone-600'}`}>
                  <Layers className="w-4 h-4 mb-1 opacity-75" />
                  <span>{slide.elements.length} OBJ</span>
                </div>

                <div
                  className={`absolute top-1 left-1 font-mono text-[10px] font-bold px-1 py-0.2 border ${
                    isSelected
                      ? 'bg-[#E4E3E0] text-[#141414] border-[#141414]'
                      : 'bg-[#141414] text-[#E4E3E0] border-[#141414]'
                  }`}
                >
                  SLD_{String(originalIndex + 1).padStart(2, '0')}
                </div>

                {slide.notes?.text && (
                  <div
                    className={`absolute top-1 right-1 p-0.5 border ${
                      isSelected
                        ? 'bg-amber-400 text-[#141414] border-[#141414]'
                        : 'bg-amber-300 text-[#141414] border-[#141414]'
                    }`}
                    title="Has Speaker Notes"
                  >
                    <FileText className="w-3 h-3" />
                  </div>
                )}
              </div>

              {/* Title & Stats */}
              <div className="flex items-center justify-between font-mono text-[11px]">
                <p className="font-bold truncate max-w-[130px]">
                  {slide.name || `Slide ${originalIndex + 1}`}
                </p>
                <span className="text-[10px] opacity-70">
                  {slide.elements.length} ELM
                </span>
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};
