import { childrenLocal, el, parseXmlRoot, serializeXml } from '../xml';
import { CT, NS } from '../xml/namespaces';
import { normalizePartPath, partExtension } from './paths';

/** `[Content_Types].xml`: default content types by extension plus per-part overrides. */
export class ContentTypes {
	readonly defaults = new Map<string, string>();
	readonly overrides = new Map<string, string>();

	static parse(xml: string): ContentTypes {
		const ct = new ContentTypes();
		const root = parseXmlRoot(xml);
		for (const d of childrenLocal(root, 'Default')) {
			if (d.attrs.Extension && d.attrs.ContentType) {
				ct.defaults.set(d.attrs.Extension.toLowerCase(), d.attrs.ContentType);
			}
		}
		for (const o of childrenLocal(root, 'Override')) {
			if (o.attrs.PartName && o.attrs.ContentType) {
				ct.overrides.set(normalizePartPath(o.attrs.PartName), o.attrs.ContentType);
			}
		}
		return ct;
	}

	/** A fresh content-types table with the defaults every package needs. */
	static createDefault(): ContentTypes {
		const ct = new ContentTypes();
		ct.defaults.set('rels', CT.relationships);
		ct.defaults.set('xml', CT.xml);
		return ct;
	}

	/** Content type of a part (override first, then default by extension). */
	get(partPath: string): string | undefined {
		const p = normalizePartPath(partPath);
		return this.overrides.get(p) ?? this.defaults.get(partExtension(p));
	}

	setOverride(partPath: string, contentType: string): void {
		this.overrides.set(normalizePartPath(partPath), contentType);
	}

	removeOverride(partPath: string): void {
		this.overrides.delete(normalizePartPath(partPath));
	}

	/** Ensure binary parts such as images have a default by extension. */
	ensureDefault(extension: string, contentType: string): void {
		const ext = extension.toLowerCase();
		if (!this.defaults.has(ext)) {
			this.defaults.set(ext, contentType);
		}
	}

	/** Part paths whose override matches a content type. */
	partsOfType(contentType: string): string[] {
		const out: string[] = [];
		for (const [p, t] of this.overrides) {
			if (t === contentType) {
				out.push(p);
			}
		}
		return out;
	}

	serialize(): string {
		const children = [
			...[...this.defaults].map(([ext, type]) =>
				el('Default', { Extension: ext, ContentType: type }),
			),
			...[...this.overrides].map(([p, type]) =>
				el('Override', { PartName: `/${p}`, ContentType: type }),
			),
		];
		return serializeXml(el('Types', { xmlns: NS.contentTypes }, children));
	}
}

/** Content type for an image by file extension. */
export function imageContentType(extension: string): string {
	switch (extension.toLowerCase()) {
		case 'png':
			return CT.png;
		case 'jpg':
		case 'jpeg':
		case 'jpe':
			return CT.jpeg;
		case 'gif':
			return CT.gif;
		case 'bmp':
			return CT.bmp;
		case 'tif':
		case 'tiff':
			return CT.tiff;
		case 'svg':
			return CT.svg;
		case 'webp':
			return CT.webp;
		case 'emf':
			return CT.emf;
		case 'wmf':
			return CT.wmf;
		default:
			return 'application/octet-stream';
	}
}
