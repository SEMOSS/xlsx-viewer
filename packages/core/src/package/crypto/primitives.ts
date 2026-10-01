/**
 * Cryptographic building blocks for ECMA-376 encryption. WebCrypto does the
 * work where it can; the plain TypeScript hashes and AES cover environments
 * without `crypto.subtle` (insecure pages) and keys WebCrypto rejects (192-bit
 * AES in Chromium).
 */
import { SoftwareAes } from './aes-soft';
import { hashSync, hmacSync, type HashAlgorithm } from './sha';

export { HASH_SIZE, type HashAlgorithm } from './sha';

const AES_BLOCK = 16;

function subtleCrypto(): SubtleCrypto | undefined {
	return globalThis.crypto?.subtle;
}

/** WebCrypto takes ArrayBuffer-backed views. */
function source(bytes: Uint8Array): Uint8Array<ArrayBuffer> {
	return bytes.buffer instanceof ArrayBuffer
		? (bytes as Uint8Array<ArrayBuffer>)
		: new Uint8Array(bytes);
}

/** Cryptographically secure random bytes. */
export function randomBytes(length: number): Uint8Array {
	const crypto = globalThis.crypto;
	if (!crypto?.getRandomValues) {
		throw new Error('No secure random number generator is available');
	}
	return crypto.getRandomValues(new Uint8Array(length));
}

export async function digest(algorithm: HashAlgorithm, data: Uint8Array): Promise<Uint8Array> {
	const subtle = subtleCrypto();
	return subtle
		? new Uint8Array(await subtle.digest(algorithm, source(data)))
		: hashSync(algorithm, data);
}

export async function hmac(
	algorithm: HashAlgorithm,
	key: Uint8Array,
	data: Uint8Array,
): Promise<Uint8Array> {
	const subtle = subtleCrypto();
	if (!subtle || key.length === 0) {
		return hmacSync(algorithm, key, data);
	}
	const hmacKey = await subtle.importKey(
		'raw',
		source(key),
		{ name: 'HMAC', hash: algorithm },
		false,
		['sign'],
	);
	return new Uint8Array(await subtle.sign('HMAC', hmacKey, source(data)));
}

/** AES in CBC mode without padding: inputs are whole 16-byte blocks. */
export interface AesCbc {
	encrypt(iv: Uint8Array, data: Uint8Array): Promise<Uint8Array>;
	decrypt(iv: Uint8Array, data: Uint8Array): Promise<Uint8Array>;
}

function webCryptoCbc(subtle: SubtleCrypto, key: CryptoKey): AesCbc {
	return {
		async encrypt(iv, data) {
			// WebCrypto pads with PKCS#7: whole-block input gets one extra block, dropped here.
			const out = await subtle.encrypt({ name: 'AES-CBC', iv: source(iv) }, key, source(data));
			return new Uint8Array(out, 0, data.length);
		},
		async decrypt(iv, data) {
			if (data.length === 0) {
				return new Uint8Array(0);
			}
			// Append a valid padding block (a full block of 0x10 chained after the last
			// ciphertext block) so the padding check passes and only that block is removed.
			const last = source(data.subarray(data.length - AES_BLOCK));
			const pad = await subtle.encrypt({ name: 'AES-CBC', iv: last }, key, new Uint8Array(0));
			const input = new Uint8Array(data.length + AES_BLOCK);
			input.set(data);
			input.set(new Uint8Array(pad, 0, AES_BLOCK), data.length);
			const plain = await subtle.decrypt({ name: 'AES-CBC', iv: source(iv) }, key, input);
			return new Uint8Array(plain);
		},
	};
}

function softwareCbc(key: Uint8Array): AesCbc {
	const aes = new SoftwareAes(key);
	return {
		encrypt: async (iv, data) => aes.encryptCbc(iv, data),
		decrypt: async (iv, data) => aes.decryptCbc(iv, data),
	};
}

/**
 * An AES-CBC cipher for a 128, 192 or 256-bit key. `software` forces the
 * TypeScript implementation (tests).
 */
export async function aesCbc(key: Uint8Array, software = false): Promise<AesCbc> {
	if (key.length !== 16 && key.length !== 24 && key.length !== 32) {
		throw new Error(`Invalid AES key length: ${key.length * 8} bits`);
	}
	const subtle = software ? undefined : subtleCrypto();
	if (subtle) {
		try {
			const usages: KeyUsage[] = ['encrypt', 'decrypt'];
			const cryptoKey = await subtle.importKey('raw', source(key), 'AES-CBC', false, usages);
			return webCryptoCbc(subtle, cryptoKey);
		} catch {
			// Not supported here (192-bit keys in Chromium): use the software cipher.
		}
	}
	return softwareCbc(key);
}

/** AES-ECB decryption, as CBC decryption with each block re-XORed with the previous one. */
export async function aesEcbDecrypt(cipher: AesCbc, data: Uint8Array): Promise<Uint8Array> {
	const out = await cipher.decrypt(new Uint8Array(AES_BLOCK), data);
	for (let i = AES_BLOCK; i < out.length; i++) {
		out[i] ^= data[i - AES_BLOCK];
	}
	return out;
}

export function concatBytes(...parts: Uint8Array[]): Uint8Array {
	const out = new Uint8Array(parts.reduce((n, p) => n + p.length, 0));
	let off = 0;
	for (const p of parts) {
		out.set(p, off);
		off += p.length;
	}
	return out;
}

export function uint32LE(value: number): Uint8Array {
	const out = new Uint8Array(4);
	new DataView(out.buffer).setUint32(0, value, true);
	return out;
}

/** Passwords are hashed as UTF-16LE code units, without a terminator. */
export function utf16le(text: string): Uint8Array {
	const out = new Uint8Array(text.length * 2);
	for (let i = 0; i < text.length; i++) {
		const c = text.charCodeAt(i);
		out[i * 2] = c & 0xff;
		out[i * 2 + 1] = c >>> 8;
	}
	return out;
}

/** Compare without stopping at the first difference. */
export function equalBytes(a: Uint8Array, b: Uint8Array): boolean {
	if (a.length !== b.length) {
		return false;
	}
	let diff = 0;
	for (let i = 0; i < a.length; i++) {
		diff |= a[i] ^ b[i];
	}
	return diff === 0;
}

/** Truncate, or pad with `fill`, to exactly `length` bytes. */
export function fitBytes(bytes: Uint8Array, length: number, fill: number): Uint8Array {
	const out = new Uint8Array(length).fill(fill);
	out.set(bytes.subarray(0, length));
	return out;
}

/** Zero-pad to a whole number of `blockSize` blocks. */
export function padToBlock(bytes: Uint8Array, blockSize: number): Uint8Array {
	const length = Math.ceil(bytes.length / blockSize) * blockSize;
	return length === bytes.length ? bytes : fitBytes(bytes, length, 0);
}

export function base64Encode(bytes: Uint8Array): string {
	let text = '';
	for (const b of bytes) {
		text += String.fromCharCode(b);
	}
	return btoa(text);
}

export function base64Decode(text: string): Uint8Array {
	return Uint8Array.from(atob(text.replace(/\s+/g, '')), (c) => c.charCodeAt(0));
}
