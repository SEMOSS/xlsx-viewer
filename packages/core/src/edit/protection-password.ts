/**
 * Passwords of sheet and workbook protection. Excel 2013+ stores a salted,
 * iterated hash (`algorithmName`, `hashValue`, `saltValue`, `spinCount`);
 * older files store a 16-bit legacy hash in `password`. Neither encrypts
 * anything: they only gate the Unprotect commands.
 */

export interface ProtectionHash {
	/** Legacy 16-bit hash as 4 hex digits. */
	password?: string;
	algorithmName?: string;
	hashValue?: string;
	saltValue?: string;
	spinCount?: number;
}

const ALGORITHMS: Record<string, string> = {
	'SHA-1': 'SHA-1',
	'SHA-256': 'SHA-256',
	'SHA-384': 'SHA-384',
	'SHA-512': 'SHA-512',
};

/** Excel's legacy protection hash (the `password` attribute). */
export function legacyPasswordHash(password: string): string {
	let hash = 0;
	for (let i = 0; i < password.length; i++) {
		const value = password.charCodeAt(i) << (i + 1);
		hash ^= (value & 0x7fff) | (value >> 15);
	}
	hash ^= password.length;
	hash ^= 0xce4b;
	return hash.toString(16).toUpperCase().padStart(4, '0');
}

function toBase64(bytes: Uint8Array): string {
	let s = '';
	for (const b of bytes) {
		s += String.fromCharCode(b);
	}
	return btoa(s);
}

function fromBase64(text: string): Uint8Array {
	const s = atob(text);
	const out = new Uint8Array(s.length);
	for (let i = 0; i < s.length; i++) {
		out[i] = s.charCodeAt(i);
	}
	return out;
}

function utf16le(text: string): Uint8Array {
	const out = new Uint8Array(text.length * 2);
	for (let i = 0; i < text.length; i++) {
		const c = text.charCodeAt(i);
		out[i * 2] = c & 0xff;
		out[i * 2 + 1] = c >> 8;
	}
	return out;
}

/** The salted hash: H(salt + password), then `spinCount` rounds of H(hash + round as uint32 LE). */
export async function iteratedPasswordHash(
	password: string,
	salt: Uint8Array,
	spinCount: number,
	algorithm = 'SHA-512',
): Promise<Uint8Array> {
	const name = ALGORITHMS[algorithm.toUpperCase()];
	if (!name) {
		throw new Error(`Unsupported protection hash algorithm: ${algorithm}`);
	}
	const subtle = globalThis.crypto.subtle;
	const pw = utf16le(password);
	const first = new Uint8Array(salt.length + pw.length);
	first.set(salt);
	first.set(pw, salt.length);
	let hash = new Uint8Array(await subtle.digest(name, first));
	const buf = new Uint8Array(hash.length + 4);
	const view = new DataView(buf.buffer);
	for (let i = 0; i < spinCount; i++) {
		buf.set(hash);
		view.setUint32(hash.length, i, true);
		hash = new Uint8Array(await subtle.digest(name, buf));
	}
	return hash;
}

/** A new SHA-512 protection hash with a random salt, as Excel writes it. */
export async function createProtectionHash(
	password: string,
	spinCount = 100_000,
): Promise<ProtectionHash> {
	const salt = globalThis.crypto.getRandomValues(new Uint8Array(16));
	const hash = await iteratedPasswordHash(password, salt, spinCount, 'SHA-512');
	return {
		algorithmName: 'SHA-512',
		hashValue: toBase64(hash),
		saltValue: toBase64(salt),
		spinCount,
	};
}

/** The password unlocks this protection (always true when it has none). */
export async function verifyProtectionPassword(
	protection: ProtectionHash,
	password: string,
): Promise<boolean> {
	if (protection.hashValue && protection.saltValue && protection.algorithmName) {
		const hash = await iteratedPasswordHash(
			password,
			fromBase64(protection.saltValue),
			protection.spinCount ?? 0,
			protection.algorithmName,
		);
		return toBase64(hash) === protection.hashValue;
	}
	if (protection.password) {
		return legacyPasswordHash(password) === protection.password.toUpperCase().padStart(4, '0');
	}
	return true;
}

/** The protection asks for a password before it can be removed. */
export function hasProtectionPassword(protection: ProtectionHash | undefined): boolean {
	return !!protection && !!(protection.hashValue || protection.password);
}
