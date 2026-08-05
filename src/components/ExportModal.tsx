import React, { useState } from 'react';
import { PresentationDocument } from '../types/pptx';
import { HtmlExporter } from '../engine/export/htmlExporter';
import { Code, Download, ExternalLink, FileCode, Check, Globe } from 'lucide-react';

interface Props {
  presentation: PresentationDocument;
  onClose: () => void;
}

export const ExportModal: React.FC<Props> = ({ presentation, onClose }) => {
  const [downloadedHtml, setDownloadedHtml] = useState(false);
  const [downloadedJson, setDownloadedJson] = useState(false);

  const handleDownloadHtml = () => {
    const htmlContent = HtmlExporter.exportStandaloneHtml(presentation);
    const blob = new Blob([htmlContent], { type: 'text/html' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${presentation.metadata.title || 'presentation'}.html`;
    a.click();
    URL.revokeObjectURL(url);
    setDownloadedHtml(true);
    setTimeout(() => setDownloadedHtml(false), 3000);
  };

  const handleDownloadJson = () => {
    const jsonContent = JSON.stringify(presentation, null, 2);
    const blob = new Blob([jsonContent], { type: 'application/json' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${presentation.metadata.title || 'presentation'}-IR.json`;
    a.click();
    URL.revokeObjectURL(url);
    setDownloadedJson(true);
    setTimeout(() => setDownloadedJson(false), 3000);
  };

  return (
    <div className="fixed inset-0 z-50 bg-[#141414]/60 backdrop-blur-xs flex items-center justify-center p-6 select-none font-mono">
      <div className="bg-[#E4E3E0] border-2 border-[#141414] max-w-lg w-full p-6 shadow-2xl flex flex-col space-y-5 text-[#141414]">
        <div className="flex items-center justify-between border-b border-[#141414] pb-3 bg-[#141414] -m-6 mb-0 p-4 text-[#E4E3E0]">
          <div className="flex items-center gap-2">
            <Globe className="w-4 h-4 text-emerald-400" />
            <h3 className="font-mono font-bold text-xs uppercase tracking-wider">// EXPORT_PACKAGE_GEN</h3>
          </div>
          <button
            onClick={onClose}
            className="text-[#141414] bg-[#E4E3E0] hover:bg-white text-xs px-2.5 py-1 font-bold border border-[#141414] transition-colors"
          >
            CLOSE
          </button>
        </div>

        <div className="space-y-4 pt-2">
          {/* Option 1: Standalone HTML */}
          <div className="bg-[#F0EFEA] border border-[#141414] p-4 space-y-2">
            <div className="flex items-center justify-between">
              <span className="font-bold text-xs uppercase text-[#141414]">// STANDALONE_HTML_PACKAGE</span>
              <span className="bg-[#141414] text-[#E4E3E0] text-[10px] px-1.5 py-0.2 font-mono">.HTML</span>
            </div>
            <p className="text-xs text-stone-700 leading-relaxed font-mono">
              Self-contained offline HTML bundle containing rendered SVG vectors, layout grids, and embedded slide deck controls.
            </p>
            <button
              onClick={handleDownloadHtml}
              className="w-full mt-2 flex items-center justify-center gap-2 bg-[#1A7F37] hover:bg-emerald-800 text-white font-bold text-xs py-2 border border-[#141414] transition-colors"
            >
              {downloadedHtml ? <Check className="w-4 h-4 text-emerald-200" /> : <Download className="w-4 h-4" />}
              <span>{downloadedHtml ? 'DOWNLOAD_COMPLETE!' : 'GENERATE_STANDALONE_HTML'}</span>
            </button>
          </div>

          {/* Option 2: Normalized IR JSON */}
          <div className="bg-[#F0EFEA] border border-[#141414] p-4 space-y-2">
            <div className="flex items-center justify-between">
              <span className="font-bold text-xs uppercase text-[#141414]">// SCENE_GRAPH_SPEC_IR</span>
              <span className="bg-[#141414] text-[#E4E3E0] text-[10px] px-1.5 py-0.2 font-mono">.JSON</span>
            </div>
            <p className="text-xs text-stone-700 leading-relaxed font-mono">
              Clean, strongly-typed OOXML Intermediate Representation JSON with normalized coordinate geometry & styling nodes.
            </p>
            <button
              onClick={handleDownloadJson}
              className="w-full mt-2 flex items-center justify-center gap-2 bg-[#141414] hover:bg-stone-800 text-[#E4E3E0] font-bold text-xs py-2 border border-[#141414] transition-colors"
            >
              {downloadedJson ? <Check className="w-4 h-4 text-emerald-300" /> : <FileCode className="w-4 h-4" />}
              <span>{downloadedJson ? 'DOWNLOAD_COMPLETE!' : 'GENERATE_SCENE_GRAPH_JSON'}</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
