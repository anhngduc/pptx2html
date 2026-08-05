import { Diagnostic } from '../../types/pptx';

export interface SecurityPolicy {
  maxArchiveBytes: number;
  maxUncompressedBytes: number;
  maxEntryCount: number;
  allowExternalUrls: boolean;
  allowedProtocols: string[];
}

export const DEFAULT_SECURITY_POLICY: SecurityPolicy = {
  maxArchiveBytes: 100 * 1024 * 1024, // 100MB
  maxUncompressedBytes: 500 * 1024 * 1024, // 500MB
  maxEntryCount: 10000,
  allowExternalUrls: true,
  allowedProtocols: ['http:', 'https:', 'mailto:', 'tel:'],
};

export function validateZipPath(path: string): string {
  // Prevent path traversal
  const normalized = path.replace(/\\/g, '/').replace(/^\/+/, '');
  if (normalized.includes('../') || normalized.includes('..\\')) {
    throw new Error(`Security Violation: Path traversal detected in zip path '${path}'`);
  }
  return normalized;
}

export function sanitizeUrl(url: string, policy = DEFAULT_SECURITY_POLICY): string | undefined {
  if (!url) return undefined;
  const trimmed = url.trim();
  if (trimmed.startsWith('#')) return trimmed; // Slide anchor reference
  
  try {
    const parsed = new URL(trimmed);
    if (policy.allowedProtocols.includes(parsed.protocol)) {
      return trimmed;
    }
  } catch {
    // Relative link
    if (!trimmed.toLowerCase().startsWith('javascript:') && !trimmed.toLowerCase().startsWith('data:text/html')) {
      return trimmed;
    }
  }
  return undefined;
}
