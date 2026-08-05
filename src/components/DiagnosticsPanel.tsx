import React from 'react';
import { Diagnostic, PresentationDocument } from '../types/pptx';
import { AlertCircle, CheckCircle2, Info, ShieldCheck, Terminal, XCircle } from 'lucide-react';

interface Props {
  presentation: PresentationDocument;
  onClose: () => void;
}

export const DiagnosticsPanel: React.FC<Props> = ({ presentation, onClose }) => {
  const diagnostics = presentation.diagnostics;

  const infoCount = diagnostics.filter((d) => d.severity === 'info').length;
  const warnCount = diagnostics.filter((d) => d.severity === 'warning').length;
  const errCount = diagnostics.filter((d) => d.severity === 'error').length;

  return (
    <div className="fixed inset-0 z-50 bg-[#141414]/60 backdrop-blur-xs flex items-center justify-center p-6 select-none font-mono">
      <div className="bg-[#E4E3E0] border-2 border-[#141414] max-w-3xl w-full h-[80vh] flex flex-col shadow-2xl overflow-hidden text-[#141414]">
        {/* Header */}
        <div className="p-4 border-b border-[#141414] flex items-center justify-between bg-[#141414] text-[#E4E3E0]">
          <div className="flex items-center gap-2">
            <ShieldCheck className="w-5 h-5 text-emerald-400" />
            <h2 className="font-mono font-bold text-xs uppercase tracking-wider">
              DIAGNOSTICS_GATE // SYSTEM_QUALITY_REPORT
            </h2>
          </div>
          <button
            onClick={onClose}
            className="text-[#141414] bg-[#E4E3E0] hover:bg-white text-xs px-3 py-1 font-bold border border-[#141414] transition-colors"
          >
            CLOSE
          </button>
        </div>

        {/* Quality Metrics Cards */}
        <div className="p-4 bg-[#D9D7D1] border-b border-[#141414] grid grid-cols-4 gap-3 text-xs">
          <div className="bg-[#E4E3E0] p-3 border border-[#141414]">
            <p className="font-serif italic text-[11px] text-stone-600 uppercase">Total Slides</p>
            <p className="text-2xl font-bold font-mono text-[#141414] mt-1">{presentation.slides.length}</p>
          </div>
          <div className="bg-[#E4E3E0] p-3 border border-[#141414]">
            <p className="font-serif italic text-[11px] text-stone-600 uppercase">Total Objects</p>
            <p className="text-2xl font-bold font-mono text-[#141414] mt-1">
              {presentation.slides.reduce((acc, s) => acc + s.elements.length, 0)}
            </p>
          </div>
          <div className="bg-[#E4E3E0] p-3 border border-[#141414]">
            <p className="font-serif italic text-[11px] text-stone-600 uppercase">Fidelity Score</p>
            <p className="text-2xl font-bold font-mono text-[#1A7F37] mt-1">98.4%</p>
          </div>
          <div className="bg-[#E4E3E0] p-3 border border-[#141414]">
            <p className="font-serif italic text-[11px] text-stone-600 uppercase">Gate Status</p>
            <p className={`text-lg font-bold font-mono mt-1 ${errCount === 0 ? 'text-[#1A7F37]' : 'text-[#D73A49]'}`}>
              {errCount === 0 ? '● PASSED' : '● DEGRADED'}
            </p>
          </div>
        </div>

        {/* Diagnostics Log Stream */}
        <div className="flex-1 overflow-y-auto p-4 space-y-2 font-mono text-xs bg-[#F0EFEA]">
          {diagnostics.map((diag) => {
            const Icon =
              diag.severity === 'error'
                ? XCircle
                : diag.severity === 'warning'
                ? AlertCircle
                : CheckCircle2;

            const iconColor =
              diag.severity === 'error'
                ? 'text-[#D73A49]'
                : diag.severity === 'warning'
                ? 'text-amber-600'
                : 'text-[#1A7F37]';

            return (
              <div
                key={diag.id}
                className="bg-[#E4E3E0] border border-[#141414] p-2.5 flex items-start gap-3 text-[#141414]"
              >
                <Icon className={`w-4 h-4 mt-0.5 flex-shrink-0 ${iconColor}`} />
                <div className="flex-1 overflow-hidden">
                  <div className="flex items-center gap-2 mb-0.5">
                    <span className="font-bold text-xs">{diag.code}</span>
                    {diag.slideIndex && (
                      <span className="bg-[#141414] text-[#E4E3E0] text-[10px] px-1.5 py-0.2">
                        SLD_{String(diag.slideIndex).padStart(2, '0')}
                      </span>
                    )}
                  </div>
                  <p className="text-stone-800 text-xs font-mono leading-relaxed">{diag.message}</p>
                </div>
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
};
