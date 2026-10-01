import { crc32 } from './crc32';
import { deflateRaw } from './streams';
import type { RawZipEntry } from './zip-read';

/** One archive entry: new contents (compressed here) or a raw entry copied unchanged. */
export type ZipWriteEntry = { name: string; data: Uint8Array } | { name: string; raw: RawZipEntry };

/** Below this size compressing does not pay off; small parts are stored as they are. */
const STORE_BELOW = 64;

const encoder = new TextEncoder();

/** 1980-01-01 00:00, as Excel writes it: identical input gives identical bytes. */
const DOS_DATE = 0x0021;
const DOS_TIME = 0;

/**
 * Assemble a ZIP archive. New entries are deflated by the platform's
 * CompressionStream; copied entries keep their compressed bytes. Archives
 * beyond 4 GB or 65535 entries are not supported (Excel's limits are lower).
 */
export async function writeZip(entries: readonly ZipWriteEntry[]): Promise<Uint8Array> {
	const prepared = await Promise.all(
		entries.map(async (e): Promise<{ name: Uint8Array; raw: RawZipEntry }> => {
			const name = encoder.encode(e.name);
			if ('raw' in e) {
				return { name, raw: e.raw };
			}
			const crc = crc32(e.data);
			if (e.data.length < STORE_BELOW) {
				return { name, raw: { method: 0, crc32: crc, size: e.data.length, compressed: e.data } };
			}
			const compressed = await deflateRaw(e.data);
			return compressed.length < e.data.length
				? { name, raw: { method: 8, crc32: crc, size: e.data.length, compressed } }
				: { name, raw: { method: 0, crc32: crc, size: e.data.length, compressed: e.data } };
		}),
	);
	let total = 22;
	for (const p of prepared) {
		total += 30 + p.name.length + p.raw.compressed.length + 46 + p.name.length;
	}
	if (total > 0xffffffff || prepared.length > 0xffff) {
		throw new RangeError('The package is too large for a ZIP archive without ZIP64');
	}
	const out = new Uint8Array(total);
	const view = new DataView(out.buffer);
	const offsets: number[] = [];
	let p = 0;
	for (const { name, raw } of prepared) {
		offsets.push(p);
		view.setUint32(p, 0x04034b50, true);
		view.setUint16(p + 4, 20, true);
		view.setUint16(p + 6, 0x800, true); // UTF-8 names
		view.setUint16(p + 8, raw.method, true);
		view.setUint16(p + 10, DOS_TIME, true);
		view.setUint16(p + 12, DOS_DATE, true);
		view.setUint32(p + 14, raw.crc32, true);
		view.setUint32(p + 18, raw.compressed.length, true);
		view.setUint32(p + 22, raw.size, true);
		view.setUint16(p + 26, name.length, true);
		view.setUint16(p + 28, 0, true);
		out.set(name, p + 30);
		out.set(raw.compressed, p + 30 + name.length);
		p += 30 + name.length + raw.compressed.length;
	}
	const dirStart = p;
	prepared.forEach(({ name, raw }, i) => {
		view.setUint32(p, 0x02014b50, true);
		view.setUint16(p + 4, 20, true);
		view.setUint16(p + 6, 20, true);
		view.setUint16(p + 8, 0x800, true);
		view.setUint16(p + 10, raw.method, true);
		view.setUint16(p + 12, DOS_TIME, true);
		view.setUint16(p + 14, DOS_DATE, true);
		view.setUint32(p + 16, raw.crc32, true);
		view.setUint32(p + 20, raw.compressed.length, true);
		view.setUint32(p + 24, raw.size, true);
		view.setUint16(p + 28, name.length, true);
		view.setUint32(p + 42, offsets[i], true);
		out.set(name, p + 46);
		p += 46 + name.length;
	});
	view.setUint32(p, 0x06054b50, true);
	view.setUint16(p + 8, prepared.length, true);
	view.setUint16(p + 10, prepared.length, true);
	view.setUint32(p + 12, p - dirStart, true);
	view.setUint32(p + 16, dirStart, true);
	return out;
}
