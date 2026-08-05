import React, { useState } from 'react';
import { PresentationDocument, SlideElement } from '../types/pptx';
import { Code, Copy, FileCode, Check, ChevronDown, ChevronRight, Search } from 'lucide-react';

interface Props {
  presentation: PresentationDocument;
  selectedElement: SlideElement | null;
  currentSlideIndex: number;
}

export const IrInspector: React.FC<Props> = ({
  presentation,
  selectedElement,
  currentSlideIndex,
}) => {
  const [activeTab, setActiveTab] = useState<'document' | 'slide' | 'element' | 'theme'>('slide');
  const [copied, setCopied] = useState(false);
  const [filterText, setFilterText] = useState('');

  const currentSlide = presentation.slides[currentSlideIndex];

  const handleCopyJson = (data: any) => {
    navigator.clipboard.writeText(JSON.stringify(data, null, 2));
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  const renderJsonTree = (data: any, depth = 0) => {
    if (typeof data !== 'object' || data === null) {
      return <span className="text-[#1A7F37] font-mono text-xs font-bold">{JSON.stringify(data)}</span>;
    }

    if (Array.isArray(data)) {
      return (
        <div className="pl-3 border-l border-[#141414]">
          <span className="text-stone-600 text-xs font-mono">[{data.length} ITEMS]</span>
          {data.map((item, idx) => (
            <div key={idx} className="my-1">
              <span className="text-stone-500 font-mono text-xs mr-2">{idx}:</span>
              {renderJsonTree(item, depth + 1)}
            </div>
          ))}
        </div>
      );
    }

    return (
      <div className="pl-3 border-l border-[#141414] space-y-1">
        {Object.entries(data).map(([key, value]) => (
          <div key={key} className="text-xs font-mono">
            <span className="text-[#141414] font-bold mr-2">{key}:</span>
            {renderJsonTree(value, depth + 1)}
          </div>
        ))}
      </div>
    );
  };

  return (
    <div className="w-80 h-full bg-[#E4E3E0] border-l border-[#141414] flex flex-col select-none text-[#141414]">
      {/* Header Tabs */}
      <div className="p-3 border-b border-[#141414] bg-[#D9D7D1] flex items-center justify-between">
        <div className="flex items-center gap-1.5 font-mono">
          <Code className="w-4 h-4 text-[#141414]" />
          <h3 className="font-bold text-xs uppercase tracking-tight">// IR_TREE_INSPECTOR</h3>
        </div>
        <button
          onClick={() =>
            handleCopyJson(
              activeTab === 'slide'
                ? currentSlide
                : activeTab === 'element'
                ? selectedElement
                : presentation
            )
          }
          className="p-1 hover:bg-[#141414] hover:text-[#E4E3E0] border border-[#141414] transition-colors"
          title="Copy JSON"
        >
          {copied ? <Check className="w-3.5 h-3.5 text-emerald-700" /> : <Copy className="w-3.5 h-3.5" />}
        </button>
      </div>

      {/* Navigation Tabs */}
      <div className="flex border-b border-[#141414] bg-[#E4E3E0] font-mono">
        <button
          onClick={() => setActiveTab('slide')}
          className={`flex-1 py-2 text-[11px] font-bold uppercase transition-colors border-r border-[#141414] ${
            activeTab === 'slide'
              ? 'bg-[#141414] text-[#E4E3E0]'
              : 'bg-[#E4E3E0] text-[#141414] hover:bg-white'
          }`}
        >
          SLD_{String(currentSlideIndex + 1).padStart(2, '0')}
        </button>
        <button
          onClick={() => setActiveTab('element')}
          disabled={!selectedElement}
          className={`flex-1 py-2 text-[11px] font-bold uppercase transition-colors border-r border-[#141414] ${
            activeTab === 'element'
              ? 'bg-[#141414] text-[#E4E3E0]'
              : 'bg-[#E4E3E0] text-[#141414] hover:bg-white disabled:opacity-40'
          }`}
        >
          ELEMENT
        </button>
        <button
          onClick={() => setActiveTab('theme')}
          className={`flex-1 py-2 text-[11px] font-bold uppercase transition-colors ${
            activeTab === 'theme'
              ? 'bg-[#141414] text-[#E4E3E0]'
              : 'bg-[#E4E3E0] text-[#141414] hover:bg-white'
          }`}
        >
          THEME
        </button>
      </div>

      {/* JSON Inspection Body */}
      <div className="flex-1 overflow-y-auto p-3 font-mono text-xs text-[#141414]">
        {activeTab === 'slide' && currentSlide && (
          <div className="space-y-3">
            <div className="bg-[#F0EFEA] p-2.5 border border-[#141414] space-y-1 text-[11px]">
              <p className="font-bold uppercase">// {currentSlide.name}</p>
              <p className="text-stone-700">Elements: {currentSlide.elements.length}</p>
              <p className="text-stone-700">Background: {currentSlide.background.type}</p>
            </div>
            {renderJsonTree(currentSlide)}
          </div>
        )}

        {activeTab === 'element' && selectedElement && (
          <div className="space-y-3">
            <div className="bg-[#F0EFEA] p-2.5 border border-[#141414] space-y-1 text-[11px]">
              <p className="font-bold text-[#141414] uppercase">// {selectedElement.name || selectedElement.type}</p>
              <p className="text-stone-700">
                Bounds (EMU): {selectedElement.bounds.xEmu}, {selectedElement.bounds.yEmu} ({selectedElement.bounds.widthEmu}x{selectedElement.bounds.heightEmu})
              </p>
              <p className="text-stone-700">Rotation: {selectedElement.transform.rotationDegrees}°</p>
              <p className="text-stone-700">XML Path: {selectedElement.source.xmlPath}</p>
            </div>
            {renderJsonTree(selectedElement)}
          </div>
        )}

        {activeTab === 'theme' && (
          <div className="space-y-3">
            <div className="bg-[#F0EFEA] p-2.5 border border-[#141414] space-y-2">
              <p className="font-bold uppercase">// {presentation.theme.name}</p>
              <div className="grid grid-cols-2 gap-2 text-[10px]">
                {Object.entries(presentation.theme.colorScheme).map(([key, color]) => (
                  <div key={key} className="flex items-center gap-1.5 bg-[#E4E3E0] p-1 border border-[#141414]">
                    <div className="w-3 h-3 border border-[#141414]" style={{ backgroundColor: color }} />
                    <span className="font-mono text-[#141414] font-bold truncate">{key}</span>
                  </div>
                ))}
              </div>
            </div>
            {renderJsonTree(presentation.theme)}
          </div>
        )}
      </div>
    </div>
  );
};
