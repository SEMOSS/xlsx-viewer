import type { OpcPackage } from '../package/opc-package';
import { resolveRelTarget } from '../package/paths';
import type { Relationship } from '../types/sheet';

let utf8Decoder: TextDecoder | undefined;

function decodeUtf16(bytes: Uint8Array, littleEndian: boolean): string {
	try {
		return new TextDecoder(littleEndian ? 'utf-16le' : 'utf-16be').decode(bytes);
	} catch {
		let out = '';
		for (let i = 0; i + 1 < bytes.length; i += 2) {
			out += String.fromCharCode(
				littleEndian ? bytes[i] | (bytes[i + 1] << 8) : (bytes[i] << 8) | bytes[i + 1],
			);
		}
		return out;
	}
}

/**
 * Decode the bytes of an XML part. UTF-8 is the norm (a UTF-8 BOM is
 * dropped); UTF-16 parts are recognised by their byte-order mark.
 */
export function decodeXmlBytes(bytes: Uint8Array): string {
	if (bytes.length >= 2) {
		if (bytes[0] === 0xff && bytes[1] === 0xfe) {
			return decodeUtf16(bytes.subarray(2), true);
		}
		if (bytes[0] === 0xfe && bytes[1] === 0xff) {
			return decodeUtf16(bytes.subarray(2), false);
		}
	}
	utf8Decoder ??= new TextDecoder('utf-8');
	return utf8Decoder.decode(bytes);
}

/**
 * Text of a part, decoded with the platform `TextDecoder` (much faster than
 * the pure-JS UTF-8 decoder JSZip falls back to in browsers).
 */
export async function readPartText(pkg: OpcPackage, partPath: string): Promise<string | undefined> {
	const bytes = await pkg.bytes(partPath);
	return bytes ? decodeXmlBytes(bytes) : undefined;
}

/** Relationship lookup by id for one source part. */
export interface RelLookup {
	readonly source: string;
	readonly rels: Relationship[];
	byId(id: string | undefined): Relationship | undefined;
	/** Package path of an internal relationship target (`undefined` for external targets). */
	path(id: string | undefined): string | undefined;
}

/**
 * Index the relationships of `source`. With `pkg`, targets written
 * percent-encoded (`image%201.png`) resolve to the part that exists.
 */
export function relLookup(source: string, rels: Relationship[], pkg?: OpcPackage): RelLookup {
	let map: Map<string, Relationship> | undefined;
	const byId = (id: string | undefined): Relationship | undefined => {
		if (!id) {
			return undefined;
		}
		if (!map) {
			map = new Map(rels.map((r) => [r.id, r]));
		}
		return map.get(id);
	};
	return {
		source,
		rels,
		byId,
		path(id) {
			const rel = byId(id);
			if (!rel || rel.targetMode === 'External') {
				return undefined;
			}
			const path = resolveRelTarget(source, rel.target);
			if (pkg && path.includes('%') && !pkg.has(path)) {
				try {
					const decoded = decodeURIComponent(path);
					return pkg.has(decoded) ? decoded : path;
				} catch {
					return path;
				}
			}
			return path;
		},
	};
}

/** Suffix of a relationship type URI (after the last `/`), which is the same in Strict and Transitional. */
export function relTypeSuffix(type: string): string {
	return type.slice(type.lastIndexOf('/') + 1);
}
