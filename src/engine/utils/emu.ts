/**
 * OOXML EMU (English Metric Units) conversion utilities.
 * 1 inch = 914,400 EMUs
 * 1 inch = 96 CSS pixels
 * 1 inch = 72 Points (pt)
 */

export const EMU_PER_INCH = 914400;
export const CSS_PX_PER_INCH = 96;
export const PT_PER_INCH = 72;

export function emuToPx(emu: number): number {
  if (!emu || isNaN(emu)) return 0;
  return (emu / EMU_PER_INCH) * CSS_PX_PER_INCH;
}

export function pxToEmu(px: number): number {
  if (!px || isNaN(px)) return 0;
  return (px / CSS_PX_PER_INCH) * EMU_PER_INCH;
}

export function emuToPt(emu: number): number {
  if (!emu || isNaN(emu)) return 0;
  return (emu / EMU_PER_INCH) * PT_PER_INCH;
}

export function ptToPx(pt: number): number {
  if (!pt || isNaN(pt)) return 0;
  return (pt / PT_PER_INCH) * CSS_PX_PER_INCH;
}

export function hundredthPtToPt(hPt: number): number {
  if (!hPt || isNaN(hPt)) return 12;
  return hPt / 100;
}

// Default 16:9 widescreen presentation slide in EMUs (1280px x 720px)
export const DEFAULT_SLIDE_WIDTH_EMU = 12192000;
export const DEFAULT_SLIDE_HEIGHT_EMU = 6858000;
