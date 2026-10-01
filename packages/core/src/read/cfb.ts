/**
 * Minimal reader for the directory of an OLE2 compound file (CFB). Encrypted
 * OOXML packages and legacy `.xls` workbooks are both compound files; the
 * stream names tell them apart (`EncryptionInfo` / `EncryptedPackage` versus
 * `Workbook` / `Book`).
 */

const MAX_REGULAR_SECTOR = 0xfffffffa;
const END_OF_CHAIN = 0xfffffffe;
const HEADER_DIFAT_ENTRIES = 109;
const DIRECTORY_ENTRY_SIZE = 128;

/**
 * Names of every directory entry (storages and streams), or `undefined` when
 * the header is not a valid compound-file header.
 */
export function compoundFileEntryNames(bytes: Uint8Array): string[] | undefined {
	if (bytes.length < 512) {
		return undefined;
	}
	const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
	const shift = view.getUint16(0x1e, true);
	if (shift !== 9 && shift !== 12) {
		return undefined;
	}
	const size = 1 << shift;
	const sectorOffset = (n: number): number => (n + 1) * size;

	// FAT sector locations: 109 in the header, the rest in the DIFAT chain.
	const fatSectors: number[] = [];
	for (let i = 0; i < HEADER_DIFAT_ENTRIES; i++) {
		const s = view.getUint32(0x4c + i * 4, true);
		if (s < MAX_REGULAR_SECTOR) {
			fatSectors.push(s);
		}
	}
	let difat = view.getUint32(0x44, true);
	for (let guard = 0; difat < MAX_REGULAR_SECTOR && guard < 1 << 16; guard++) {
		const off = sectorOffset(difat);
		if (off + size > bytes.length) {
			break;
		}
		for (let i = 0; i < size / 4 - 1; i++) {
			const s = view.getUint32(off + i * 4, true);
			if (s < MAX_REGULAR_SECTOR) {
				fatSectors.push(s);
			}
		}
		difat = view.getUint32(off + size - 4, true);
	}
	const perSector = size / 4;
	const nextSector = (n: number): number => {
		const fat = fatSectors[Math.floor(n / perSector)];
		const off = fat === undefined ? -1 : sectorOffset(fat) + (n % perSector) * 4;
		return off >= 0 && off + 4 <= bytes.length ? view.getUint32(off, true) : END_OF_CHAIN;
	};

	const names: string[] = [];
	let sector = view.getUint32(0x30, true);
	for (let guard = 0; sector < MAX_REGULAR_SECTOR && guard < 1 << 20; guard++) {
		const off = sectorOffset(sector);
		if (off + size > bytes.length) {
			break;
		}
		for (let e = off; e < off + size; e += DIRECTORY_ENTRY_SIZE) {
			const len = view.getUint16(e + 0x40, true);
			if (len < 2 || len > 64 || bytes[e + 0x42] === 0) {
				continue;
			}
			let name = '';
			for (let i = 0; i < len - 2; i += 2) {
				name += String.fromCharCode(view.getUint16(e + i, true));
			}
			names.push(name);
		}
		sector = nextSector(sector);
	}
	return names;
}
