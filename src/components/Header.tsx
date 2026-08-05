import React, { useRef } from 'react';
import { AVAILABLE_SAMPLES } from '../engine/sample/sampleGenerator';
import {
  Code,
  Download,
  FileSpreadsheet,
  FolderOpen,
  Globe,
  Play,
  Presentation,
  ShieldCheck,
  Upload,
} from 'lucide-react';

interface Props {
  title: string;
  onLoadSample: (sampleId: string) => void;
  onFileUpload: (file: File) => void;
  onTogglePresenter: () => void;
  onToggleDiagnostics: () => void;
  onToggleInspector: () => void;
  onOpenExport: () => void;
  inspectorOpen: boolean;
  isLoading: boolean;
}

export const Header: React.FC<Props> = ({
  title,
  onLoadSample,
  onFileUpload,
  onTogglePresenter,
  onToggleDiagnostics,
  onToggleInspector,
  onOpenExport,
  inspectorOpen,
  isLoading,
}) => {
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) {
      onFileUpload(file);
    }
  };

  return (
    <header className="h-14 bg-[#E4E3E0] border-b border-[#141414] px-4 flex items-center justify-between select-none z-30 text-[#141414]">
      {/* Brand Title */}
      <div className="flex items-center gap-3">
        <div className="w-7 h-7 bg-[#141414] text-[#E4E3E0] flex items-center justify-center font-mono text-xs font-bold shadow-none">
          <Presentation className="w-4 h-4" />
        </div>
        <div className="flex flex-col">
          <div className="flex items-center gap-2">
            <h1 className="font-mono font-bold text-xs tracking-tight text-[#141414] uppercase">
              PPTX_RENDERER_SYSTEM / v2.4.0
            </h1>
            <span className="font-serif italic text-[11px] text-stone-600 hidden sm:inline">
              [SYSTEM_OK]
            </span>
          </div>
          <p className="text-[10px] font-mono text-stone-600 truncate max-w-[200px] sm:max-w-xs">
            {title ? `FILE: ${title}` : 'IDLE / READY'}
          </p>
        </div>
      </div>

      {/* Action Controls & Sample Picker */}
      <div className="flex items-center gap-2 text-xs font-mono">
        {/* Sample Dropdown */}
        <div className="relative group">
          <button className="flex items-center gap-1.5 bg-[#E4E3E0] hover:bg-[#141414] hover:text-[#E4E3E0] text-[#141414] text-xs px-2.5 py-1.5 border border-[#141414] transition-colors">
            <FolderOpen className="w-3.5 h-3.5" />
            <span className="hidden sm:inline">LOAD_SAMPLE</span>
          </button>
          <div className="absolute right-0 top-full mt-1 w-64 bg-[#E4E3E0] border border-[#141414] shadow-2xl p-2 hidden group-hover:block z-50">
            <p className="text-[10px] font-mono font-bold text-stone-600 uppercase tracking-wider px-2 py-1 mb-1 border-b border-[#141414]/30">
              // SYNTHETIC OOXML SAMPLES
            </p>
            {AVAILABLE_SAMPLES.map((sample) => (
              <button
                key={sample.id}
                onClick={() => onLoadSample(sample.id)}
                className="w-full text-left p-2 hover:bg-[#141414] hover:text-[#E4E3E0] transition-colors flex flex-col gap-0.5 border border-transparent hover:border-[#141414]"
              >
                <span className="text-xs font-bold">{sample.title}</span>
                <span className="text-[10px] opacity-75 line-clamp-1">{sample.description}</span>
              </button>
            ))}
          </div>
        </div>

        {/* Upload Button */}
        <input
          type="file"
          ref={fileInputRef}
          accept=".pptx"
          onChange={handleFileChange}
          className="hidden"
        />
        <button
          onClick={() => fileInputRef.current?.click()}
          disabled={isLoading}
          className="flex items-center gap-1.5 bg-[#141414] hover:bg-stone-800 disabled:opacity-50 text-[#E4E3E0] text-xs font-mono font-bold px-3 py-1.5 border border-[#141414] transition-colors"
        >
          <Upload className="w-3.5 h-3.5" />
          <span>UPLOAD .PPTX</span>
        </button>

        <div className="h-5 w-[1px] bg-[#141414] my-auto mx-1 opacity-40" />

        {/* IR Inspector Toggle */}
        <button
          onClick={onToggleInspector}
          className={`flex items-center gap-1.5 text-xs px-2.5 py-1.5 border border-[#141414] transition-colors ${
            inspectorOpen
              ? 'bg-[#141414] text-[#E4E3E0]'
              : 'bg-[#E4E3E0] hover:bg-[#141414] hover:text-[#E4E3E0] text-[#141414]'
          }`}
          title="Toggle IR Tree Inspector"
        >
          <Code className="w-3.5 h-3.5" />
          <span className="hidden md:inline">IR_INSPECTOR</span>
        </button>

        {/* Diagnostics & Quality Gate */}
        <button
          onClick={onToggleDiagnostics}
          className="flex items-center gap-1.5 bg-[#E4E3E0] hover:bg-[#141414] hover:text-[#E4E3E0] border border-[#141414] text-[#141414] text-xs px-2.5 py-1.5 transition-colors"
          title="Diagnostics & Visual Gate"
        >
          <ShieldCheck className="w-3.5 h-3.5 text-emerald-700 group-hover:text-emerald-400" />
          <span className="hidden md:inline">QUALITY_GATE</span>
        </button>

        {/* Presenter Mode */}
        <button
          onClick={onTogglePresenter}
          className="flex items-center gap-1.5 bg-[#E4E3E0] hover:bg-[#141414] hover:text-[#E4E3E0] border border-[#141414] text-[#141414] text-xs px-2.5 py-1.5 transition-colors"
          title="Presenter View"
        >
          <Play className="w-3.5 h-3.5" />
          <span className="hidden lg:inline">PRESENTER</span>
        </button>

        {/* Export HTML */}
        <button
          onClick={onOpenExport}
          className="flex items-center gap-1.5 bg-[#1A7F37] hover:bg-emerald-800 text-white text-xs font-mono font-bold px-3 py-1.5 border border-[#141414] transition-colors"
        >
          <Globe className="w-3.5 h-3.5" />
          <span>EXPORT_HTML</span>
        </button>
      </div>
    </header>
  );
};
