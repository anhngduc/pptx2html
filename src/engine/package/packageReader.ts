import JSZip from 'jszip';
import { validateZipPath } from './security';
import { ResourceStore } from '../../types/pptx';
import { XmlUtils } from '../parser/xmlUtils';

export interface Relationship {
  id: string;
  type: string;
  target: string;
  targetMode?: 'Internal' | 'External';
}

export interface PackagePart {
  path: string;
  contentType: string;
  textData?: string;
  binaryData?: Uint8Array;
}

export class OoxmlPackageReader {
  private zip: JSZip | null = null;
  private parts = new Map<string, PackagePart>();
  private relationships = new Map<string, Map<string, Relationship>>();
  public objectUrlsToRevoke: string[] = [];

  async load(source: File | Blob | ArrayBuffer | Uint8Array): Promise<ResourceStore> {
    this.zip = await JSZip.loadAsync(source);
    this.parts.clear();
    this.relationships.clear();

    const resourceStore: ResourceStore = {
      images: new Map(),
      media: new Map(),
      embeddedFiles: new Map(),
    };

    // Index all entries
    const entries = Object.keys(this.zip.files);
    for (const filename of entries) {
      const zipObj = this.zip.files[filename];
      if (zipObj.dir) continue;

      const safePath = validateZipPath(filename);
      if (safePath.endsWith('.rels')) {
        await this.parseRelationships(safePath, zipObj);
      } else if (
        safePath.endsWith('.xml') ||
        safePath.endsWith('.rels') ||
        safePath.endsWith('.txt') ||
        safePath.endsWith('.json')
      ) {
        const textData = await zipObj.async('text');
        this.parts.set(safePath, {
          path: safePath,
          contentType: this.guessContentType(safePath),
          textData,
        });
      } else {
        // Binary asset (image, media, font, chart workbook, OLE)
        const binaryData = await zipObj.async('uint8array');
        const mimeType = this.guessContentType(safePath);
        
        // Create Blob Object URL for browser rendering
        const blob = new Blob([binaryData], { type: mimeType });
        const objectUrl = URL.createObjectURL(blob);
        this.objectUrlsToRevoke.push(objectUrl);

        this.parts.set(safePath, {
          path: safePath,
          contentType: mimeType,
          binaryData,
        });

        if (mimeType.startsWith('image/')) {
          resourceStore.images.set(safePath, { mimeType, url: objectUrl, data: binaryData });
        } else if (mimeType.startsWith('video/') || mimeType.startsWith('audio/')) {
          resourceStore.media.set(safePath, { mimeType, url: objectUrl });
        } else {
          resourceStore.embeddedFiles.set(safePath, { fileName: safePath.split('/').pop() || 'embedded', data: binaryData });
        }
      }
    }

    return resourceStore;
  }

  private async parseRelationships(relsPath: string, zipObj: JSZip.JSZipObject): Promise<void> {
    const xmlText = await zipObj.async('text');
    const xmlDoc = XmlUtils.parseXml(xmlText);
    const rels = new Map<string, Relationship>();

    const relNodes = xmlDoc.getElementsByTagName('Relationship');
    for (let i = 0; i < relNodes.length; i++) {
      const node = relNodes[i];
      const id = node.getAttribute('Id') || '';
      const type = node.getAttribute('Type') || '';
      const target = node.getAttribute('Target') || '';
      const targetMode = node.getAttribute('TargetMode') as 'Internal' | 'External' | null;

      if (id && target) {
        rels.set(id, {
          id,
          type,
          target,
          targetMode: targetMode || 'Internal',
        });
      }
    }

    // Key relationships by source part path
    // e.g. 'ppt/_rels/presentation.xml.rels' -> source is 'ppt/presentation.xml'
    const sourcePath = relsPath
      .replace(/_rels\//, '')
      .replace(/\.rels$/, '');

    this.relationships.set(sourcePath, rels);
  }

  getPart(path: string): PackagePart | undefined {
    const clean = path.replace(/^\//, '');
    return this.parts.get(clean);
  }

  getRelationships(sourcePath: string): Map<string, Relationship> | undefined {
    const clean = sourcePath.replace(/^\//, '');
    return this.relationships.get(clean);
  }

  resolveTarget(sourcePath: string, rId: string): { path?: string; relationship?: Relationship } {
    const rels = this.getRelationships(sourcePath);
    if (!rels) return {};

    const rel = rels.get(rId);
    if (!rel) return {};

    if (rel.targetMode === 'External') {
      return { relationship: rel };
    }

    // Resolve relative OOXML path
    const sourceDir = sourcePath.includes('/') ? sourcePath.substring(0, sourcePath.lastIndexOf('/')) : '';
    let targetPath = rel.target;

    if (targetPath.startsWith('/')) {
      targetPath = targetPath.substring(1);
    } else if (sourceDir) {
      const parts = sourceDir.split('/');
      const targetParts = targetPath.split('/');
      for (const p of targetParts) {
        if (p === '..') {
          parts.pop();
        } else if (p !== '.') {
          parts.push(p);
        }
      }
      targetPath = parts.join('/');
    }

    return { path: targetPath, relationship: rel };
  }

  private guessContentType(path: string): string {
    const ext = path.substring(path.lastIndexOf('.')).toLowerCase();
    switch (ext) {
      case '.xml': return 'application/xml';
      case '.png': return 'image/png';
      case '.jpg':
      case '.jpeg': return 'image/jpeg';
      case '.gif': return 'image/gif';
      case '.svg': return 'image/svg+xml';
      case '.mp4': return 'video/mp4';
      case '.mp3': return 'audio/mpeg';
      case '.wav': return 'audio/wav';
      default: return 'application/octet-stream';
    }
  }

  dispose(): void {
    for (const url of this.objectUrlsToRevoke) {
      URL.revokeObjectURL(url);
    }
    this.objectUrlsToRevoke = [];
    this.parts.clear();
    this.relationships.clear();
  }
}
