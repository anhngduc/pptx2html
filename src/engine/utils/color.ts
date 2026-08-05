import { ColorScheme } from '../../types/pptx';

export const DEFAULT_THEME_COLORS: ColorScheme = {
  dk1: '#000000',
  lt1: '#FFFFFF',
  dk2: '#1F2937',
  lt2: '#F3F4F6',
  accent1: '#2563EB', // Blue
  accent2: '#059669', // Emerald
  accent3: '#D97706', // Amber
  accent4: '#DC2626', // Red
  accent5: '#7C3AED', // Purple
  accent6: '#0891B2', // Cyan
  hlink: '#2563EB',
  folHlink: '#7C3AED',
};

/**
 * Apply luminance modulation, luminance offset, tint, or shade to a hex color
 */
export function applyColorModifiers(
  hexColor: string,
  modifiers: {
    lumMod?: number; // 0 to 100000 (percentage)
    lumOff?: number; // 0 to 100000 (percentage)
    tint?: number; // 0 to 100000
    shade?: number; // 0 to 100000
    alpha?: number; // 0 to 100000
  }
): string {
  if (!hexColor || !hexColor.startsWith('#')) return hexColor || '#000000';

  let hex = hexColor.replace('#', '');
  if (hex.length === 3) {
    hex = hex.split('').map((c) => c + c).join('');
  }

  let r = parseInt(hex.substring(0, 2), 16) || 0;
  let g = parseInt(hex.substring(2, 4), 16) || 0;
  let b = parseInt(hex.substring(4, 6), 16) || 0;

  // Apply tint (mix with white)
  if (modifiers.tint !== undefined) {
    const factor = modifiers.tint / 100000;
    r = Math.round(r * factor + 255 * (1 - factor));
    g = Math.round(g * factor + 255 * (1 - factor));
    b = Math.round(b * factor + 255 * (1 - factor));
  }

  // Apply shade (mix with black)
  if (modifiers.shade !== undefined) {
    const factor = modifiers.shade / 100000;
    r = Math.round(r * factor);
    g = Math.round(g * factor);
    b = Math.round(b * factor);
  }

  // Apply luminance mod/off
  if (modifiers.lumMod !== undefined || modifiers.lumOff !== undefined) {
    const mod = (modifiers.lumMod ?? 100000) / 100000;
    const off = (modifiers.lumOff ?? 0) / 100000;
    r = Math.min(255, Math.max(0, Math.round(r * mod + 255 * off)));
    g = Math.min(255, Math.max(0, Math.round(g * mod + 255 * off)));
    b = Math.min(255, Math.max(0, Math.round(b * mod + 255 * off)));
  }

  const alpha = modifiers.alpha !== undefined ? Math.max(0, Math.min(1, modifiers.alpha / 100000)) : 1;

  if (alpha < 1) {
    return `rgba(${r}, ${g}, ${b}, ${alpha.toFixed(2)})`;
  }

  const toHex = (n: number) => n.toString(16).padStart(2, '0').toUpperCase();
  return `#${toHex(r)}${toHex(g)}${toHex(b)}`;
}

export function resolveSchemeColor(schemeName: string, themeColors: ColorScheme): string {
  const name = schemeName.toLowerCase();
  if (name in themeColors) {
    return themeColors[name];
  }
  // Standard mappings
  if (name === 'bg1') return themeColors.lt1 || '#FFFFFF';
  if (name === 'tx1') return themeColors.dk1 || '#000000';
  if (name === 'bg2') return themeColors.lt2 || '#F3F4F6';
  if (name === 'tx2') return themeColors.dk2 || '#1F2937';
  
  return themeColors.accent1 || '#2563EB';
}
