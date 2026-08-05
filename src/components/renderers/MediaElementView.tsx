import React from 'react';
import { ConnectorElement, ImageElement, MediaElement, OleElement, UnknownElement } from '../../types/pptx';
import { Download, FileText, Film, Image as ImageIcon, Volume2 } from 'lucide-react';

export const ImageElementView: React.FC<{ element: ImageElement }> = ({ element }) => {
  if (!element.srcUrl) {
    return (
      <div className="w-full h-full bg-slate-800 border border-slate-700 rounded flex flex-col items-center justify-center text-slate-400 p-2 text-xs">
        <ImageIcon className="w-6 h-6 mb-1 text-slate-500" />
        <span>Image Placeholder</span>
      </div>
    );
  }

  return (
    <img
      src={element.srcUrl}
      alt={element.altText || element.name || 'Slide Picture'}
      className="w-full h-full object-contain"
    />
  );
};

export const MediaElementView: React.FC<{ element: MediaElement }> = ({ element }) => {
  if (element.mediaKind === 'video') {
    return (
      <div className="w-full h-full bg-black rounded overflow-hidden relative flex items-center justify-center">
        {element.srcUrl ? (
          <video src={element.srcUrl} controls autoPlay={element.autoPlay} loop={element.loop} className="w-full h-full" />
        ) : (
          <div className="text-slate-400 flex flex-col items-center text-xs">
            <Film className="w-8 h-8 mb-1 text-blue-400" />
            <span>Embedded Video ({element.name})</span>
          </div>
        )}
      </div>
    );
  }

  return (
    <div className="w-full h-full bg-slate-800 rounded p-2 flex items-center gap-2 text-slate-200 text-xs">
      <Volume2 className="w-5 h-5 text-emerald-400" />
      <span>Audio: {element.name}</span>
    </div>
  );
};

export const OleElementView: React.FC<{ element: OleElement }> = ({ element }) => {
  const handleDownload = () => {
    if (element.fileData) {
      const blob = new Blob([element.fileData], { type: 'application/octet-stream' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = element.fileName || 'embedded-file.bin';
      a.click();
      URL.revokeObjectURL(url);
    }
  };

  return (
    <div className="w-full h-full bg-slate-800/90 border border-slate-700 rounded-lg p-3 flex flex-col justify-between text-slate-200">
      <div className="flex items-center gap-2">
        <FileText className="w-5 h-5 text-amber-400" />
        <div className="overflow-hidden">
          <p className="font-semibold text-xs truncate">{element.fileName || 'Embedded OLE Document'}</p>
          <p className="text-[10px] text-slate-400">{element.programId || 'Office Binary Part'}</p>
        </div>
      </div>
      <button
        onClick={handleDownload}
        className="mt-2 flex items-center justify-center gap-1.5 bg-slate-700 hover:bg-slate-600 text-white text-xs px-2.5 py-1.5 rounded transition"
      >
        <Download className="w-3.5 h-3.5" />
        <span>Extract Embedded Asset</span>
      </button>
    </div>
  );
};

export const ConnectorElementView: React.FC<{ element: ConnectorElement }> = ({ element }) => {
  return (
    <svg className="w-full h-full overflow-visible pointer-events-none">
      <line
        x1={0}
        y1={0}
        x2="100%"
        y2="100%"
        stroke={element.line.color}
        strokeWidth={element.line.widthPx}
        strokeDasharray={element.line.dash === 'dashed' ? '4,4' : undefined}
      />
    </svg>
  );
};

export const UnknownElementView: React.FC<{ element: UnknownElement }> = ({ element }) => {
  return (
    <div className="w-full h-full border border-dashed border-amber-500/60 bg-amber-500/10 rounded p-2 flex items-center justify-center text-amber-300 text-xs">
      <span>{element.fallbackText || `Unsupported OOXML Part (${element.rawXmlType})`}</span>
    </div>
  );
};
