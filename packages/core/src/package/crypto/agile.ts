/**
 * ECMA-376 Agile encryption ([MS-OFFCRYPTO] 2.3.4.10 to 2.3.4.15), used by
 * Excel 2010 and later: password key derivation, the password verifier, the
 * intermediate (secret) key, decryption of the 4096-byte package segments and
 * the data integrity HMAC. Writing lives in `agile-encrypt.ts`.
 *
 * Adapted from pptx-viewer (https://github.com/ChristopherVR/pptx-viewer), Apache-2.0.
 */
import type { AgileCipherParams, AgileEncryptionInfo, AgilePasswordKey } from './encryption-info';
import { damagedPackage, integrityCheckFailed, unsupportedEncryption } from './errors';
import {
	HASH_SIZE,
	aesCbc,
	concatBytes,
	digest,
	equalBytes,
	fitBytes,
	hmac,
	uint32LE,
	utf16le,
	type AesCbc,
	type HashAlgorithm,
} from './primitives';
import { hashSync } from './sha';

export const SEGMENT_SIZE = 4096;
/** The largest spin count [MS-OFFCRYPTO] allows. */
const MAX_SPIN_COUNT = 10_000_000;

/** Block keys of the agile key derivation ([MS-OFFCRYPTO] 2.3.4.13 and 2.3.4.14). */
export const BLOCK_KEYS = {
	verifierHashInput: Uint8Array.of(0xfe, 0xa7, 0xd2, 0x76, 0x3b, 0x4b, 0x9e, 0x79),
	verifierHashValue: Uint8Array.of(0xd7, 0xaa, 0x0f, 0x6d, 0x30, 0x61, 0x34, 0x4e),
	encryptedKeyValue: Uint8Array.of(0x14, 0x6e, 0x0b, 0xe7, 0xab, 0xac, 0xd0, 0xd6),
	hmacKey: Uint8Array.of(0x5f, 0xb2, 0xad, 0x01, 0x0c, 0xb9, 0xe1, 0xf6),
	hmacValue: Uint8Array.of(0xa0, 0x67, 0x7f, 0x02, 0xb2, 0x2c, 0x84, 0x33),
};

/** The hash a descriptor names (`SHA512`, `SHA-1`, ...), when supported. */
export function agileHashAlgorithm(name: string): HashAlgorithm | undefined {
	const names: Record<string, HashAlgorithm> = {
		SHA1: 'SHA-1',
		SHA256: 'SHA-256',
		SHA384: 'SHA-384',
		SHA512: 'SHA-512',
	};
	return names[name.replace(/-/g, '').toUpperCase()];
}

/** Check that the parameters are supported (AES-CBC, SHA-1/2); returns the hash. */
export function checkCipherParams(p: AgileCipherParams): HashAlgorithm {
	if (p.cipherAlgorithm.toUpperCase() !== 'AES') {
		throw unsupportedEncryption(`the ${p.cipherAlgorithm || 'unnamed'} cipher`);
	}
	if (p.cipherChaining.toLowerCase() !== 'chainingmodecbc') {
		throw unsupportedEncryption(`${p.cipherChaining || 'unnamed'} block chaining`);
	}
	if (p.keyBits !== 128 && p.keyBits !== 192 && p.keyBits !== 256) {
		throw unsupportedEncryption(`${p.keyBits}-bit AES keys`);
	}
	if (p.blockSize !== 16) {
		throw unsupportedEncryption(`a ${p.blockSize}-byte cipher block`);
	}
	const hash = agileHashAlgorithm(p.hashAlgorithm);
	if (!hash) {
		throw unsupportedEncryption(`the ${p.hashAlgorithm || 'unnamed'} hash algorithm`);
	}
	if (p.hashSize !== HASH_SIZE[hash] || p.saltValue.length === 0) {
		throw damagedPackage(`the ${p.hashAlgorithm} key parameters are inconsistent`);
	}
	return hash;
}

/**
 * The iterated password hash ([MS-OFFCRYPTO] 2.3.4.11): H(salt + password),
 * then `spinCount` rounds of H(iterator + previous hash), with the
 * synchronous hashes (an awaited WebCrypto call per round is far slower).
 */
export function spinPasswordHash(
	hash: HashAlgorithm,
	salt: Uint8Array,
	password: string,
	spinCount: number,
): Uint8Array {
	let h = hashSync(hash, concatBytes(salt, utf16le(password)));
	const round = new Uint8Array(4 + h.length);
	const view = new DataView(round.buffer);
	for (let i = 0; i < spinCount; i++) {
		view.setUint32(0, i, true);
		round.set(h, 4);
		h = hashSync(hash, round);
	}
	return h;
}

/** Key for a block key: H(spun hash + block key), truncated or padded with 0x36. */
export async function blockKeyCipher(
	hash: HashAlgorithm,
	spun: Uint8Array,
	blockKey: Uint8Array,
	keyBits: number,
): Promise<AesCbc> {
	return aesCbc(fitBytes(await digest(hash, concatBytes(spun, blockKey)), keyBits / 8, 0x36));
}

/** IV from a salt and block key: H(salt + block key), truncated or padded with 0x36. */
export async function blockIv(
	p: AgileCipherParams,
	hash: HashAlgorithm,
	blockKey: Uint8Array,
): Promise<Uint8Array> {
	return fitBytes(await digest(hash, concatBytes(p.saltValue, blockKey)), p.blockSize, 0x36);
}

function wholeBlocks(bytes: Uint8Array, minimum: number, what: string): Uint8Array {
	if (bytes.length < minimum || bytes.length % 16 !== 0) {
		throw damagedPackage(`the ${what} has the wrong length`);
	}
	return bytes;
}

/** Throw when the file cannot be decrypted with a password here; returns the password key. */
export function checkAgileSupport(info: AgileEncryptionInfo): AgilePasswordKey {
	const key = info.passwordKey;
	if (!key) {
		throw unsupportedEncryption('it can only be opened with a certificate');
	}
	checkCipherParams(key);
	checkCipherParams(info.keyData);
	if (key.spinCount > MAX_SPIN_COUNT) {
		throw unsupportedEncryption(`a spin count of ${key.spinCount}`);
	}
	return key;
}

/**
 * Verify the password and decrypt the intermediate key the package is
 * encrypted with. Returns `undefined` when the password is wrong.
 */
export async function agileSecretKey(
	info: AgileEncryptionInfo,
	password: string,
): Promise<Uint8Array | undefined> {
	const key = checkAgileSupport(info);
	const hash = checkCipherParams(key);
	const spun = spinPasswordHash(hash, key.saltValue, password, key.spinCount);
	// The password key encryptor uses its salt as the IV.
	const iv = fitBytes(key.saltValue, key.blockSize, 0x36);
	const decrypt = async (blockKey: Uint8Array, data: Uint8Array): Promise<Uint8Array> =>
		(await blockKeyCipher(hash, spun, blockKey, key.keyBits)).decrypt(iv, data);
	const verifier = wholeBlocks(key.encryptedVerifierHashInput, key.saltSize, 'verifier');
	const verifierHash = wholeBlocks(key.encryptedVerifierHashValue, key.hashSize, 'verifier hash');
	const [input, value] = await Promise.all([
		decrypt(BLOCK_KEYS.verifierHashInput, verifier),
		decrypt(BLOCK_KEYS.verifierHashValue, verifierHash),
	]);
	const expected = await digest(hash, input.subarray(0, key.saltSize));
	if (!equalBytes(expected, value.subarray(0, key.hashSize))) {
		return undefined;
	}
	const secretBytes = info.keyData.keyBits / 8;
	const encryptedKey = wholeBlocks(key.encryptedKeyValue, secretBytes, 'encrypted key');
	return (await decrypt(BLOCK_KEYS.encryptedKeyValue, encryptedKey)).slice(0, secretBytes);
}

/** Check the HMAC over the whole EncryptedPackage stream ([MS-OFFCRYPTO] 2.3.4.14). */
export async function checkAgileIntegrity(
	info: AgileEncryptionInfo,
	secretKey: Uint8Array,
	encryptedPackage: Uint8Array,
): Promise<void> {
	const integrity = info.dataIntegrity;
	if (!integrity) {
		return;
	}
	const kd = info.keyData;
	const hash = checkCipherParams(kd);
	const cipher = await aesCbc(secretKey);
	const [hmacKey, expected] = await Promise.all([
		blockIv(kd, hash, BLOCK_KEYS.hmacKey).then((iv) =>
			cipher.decrypt(iv, wholeBlocks(integrity.encryptedHmacKey, kd.hashSize, 'HMAC key')),
		),
		blockIv(kd, hash, BLOCK_KEYS.hmacValue).then((iv) =>
			cipher.decrypt(iv, wholeBlocks(integrity.encryptedHmacValue, kd.hashSize, 'HMAC value')),
		),
	]);
	const actual = await hmac(hash, hmacKey.subarray(0, kd.hashSize), encryptedPackage);
	if (!equalBytes(actual, expected.subarray(0, kd.hashSize))) {
		throw integrityCheckFailed();
	}
}

/** The unencrypted size in the first 8 bytes of an EncryptedPackage stream. */
export function encryptedPackageSize(stream: Uint8Array): number {
	if (stream.length < 8) {
		throw damagedPackage('the EncryptedPackage stream is truncated');
	}
	const view = new DataView(stream.buffer, stream.byteOffset, 8);
	const low = view.getUint32(0, true);
	const size = low + view.getUint32(4, true) * 0x100000000;
	// Some writers leave garbage in the high word: fall back to the low one.
	const fit = size <= stream.length - 8 ? size : low;
	if (fit > stream.length - 8) {
		throw damagedPackage('the EncryptedPackage stream is truncated');
	}
	return fit;
}

/** Decrypt the package: each 4096-byte segment is CBC-encrypted with IV H(salt + segment). */
export async function decryptAgilePackage(
	info: AgileEncryptionInfo,
	secretKey: Uint8Array,
	encryptedPackage: Uint8Array,
): Promise<Uint8Array> {
	const kd = info.keyData;
	const hash = checkCipherParams(kd);
	const size = encryptedPackageSize(encryptedPackage);
	const data = encryptedPackage.subarray(8);
	if (data.length % kd.blockSize !== 0) {
		throw damagedPackage('the encrypted data is not a whole number of cipher blocks');
	}
	if (data.length === 0) {
		return new Uint8Array(0);
	}
	const segments = Math.ceil(data.length / SEGMENT_SIZE);
	const ivs = await Promise.all(
		Array.from({ length: segments }, (_, i) => blockIv(kd, hash, uint32LE(i))),
	);
	// One CBC pass decrypts every segment as if chained to the one before;
	// re-chain the first block of each segment to its own IV instead.
	const out = await (await aesCbc(secretKey)).decrypt(ivs[0], data);
	for (let s = 1; s < segments; s++) {
		const off = s * SEGMENT_SIZE;
		for (let i = 0; i < kd.blockSize; i++) {
			out[off + i] ^= data[off - kd.blockSize + i] ^ ivs[s][i];
		}
	}
	return out.subarray(0, size);
}
