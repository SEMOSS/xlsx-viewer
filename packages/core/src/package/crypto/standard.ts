/**
 * ECMA-376 Standard encryption ([MS-OFFCRYPTO] 2.3.4.5 to 2.3.4.9), written by
 * Excel 2007: AES-ECB with a key derived from 50000 rounds of SHA-1 over the
 * salt and password.
 *
 * Adapted from pptx-viewer (https://github.com/ChristopherVR/pptx-viewer), Apache-2.0.
 * Unlike the original, the key is always expanded through X1 / X2 and the data
 * is decrypted as ECB, as [MS-OFFCRYPTO] specifies.
 */
import { encryptedPackageSize, spinPasswordHash } from './agile';
import type { StandardEncryptionInfo } from './encryption-info';
import { damagedPackage, unsupportedEncryption } from './errors';
import { aesCbc, aesEcbDecrypt, concatBytes, digest, equalBytes, uint32LE } from './primitives';

const FLAG_CRYPTO_API = 0x04;
const FLAG_AES = 0x20;
const ALG_RC4 = 0x6801;
const ALG_SHA1 = 0x8004;
const AES_KEY_BITS: Readonly<Record<number, number>> = { 0x660e: 128, 0x660f: 192, 0x6610: 256 };
const SPIN_COUNT = 50_000;
const SHA1_SIZE = 20;

/** Key size of a supported header (CryptoAPI AES with SHA-1); throws otherwise. */
export function standardKeyBits(info: StandardEncryptionInfo): number {
	const aes = (info.flags & FLAG_AES) !== 0;
	if (info.algId === ALG_RC4 || (info.algId === 0 && !aes)) {
		throw unsupportedEncryption('RC4 encryption');
	}
	const bits = info.algId === 0 ? 128 : AES_KEY_BITS[info.algId];
	if (!bits || !aes || !(info.flags & FLAG_CRYPTO_API)) {
		throw unsupportedEncryption(`the CryptoAPI algorithm 0x${info.algId.toString(16)}`);
	}
	if (info.algIdHash !== 0 && info.algIdHash !== ALG_SHA1) {
		throw unsupportedEncryption(`the CryptoAPI hash 0x${info.algIdHash.toString(16)}`);
	}
	return bits;
}

/**
 * The AES key for a password ([MS-OFFCRYPTO] 2.3.4.7): the spun SHA-1 hash,
 * finished with block 0, then expanded as X1 = H(0x36 pad), X2 = H(0x5C pad).
 */
export async function standardKey(
	info: StandardEncryptionInfo,
	password: string,
): Promise<Uint8Array> {
	const keyBits = standardKeyBits(info);
	const spun = spinPasswordHash('SHA-1', info.salt, password, SPIN_COUNT);
	const final = await digest('SHA-1', concatBytes(spun, uint32LE(0)));
	const expand = (fill: number): Promise<Uint8Array> => {
		const buffer = new Uint8Array(64).fill(fill);
		final.forEach((b, i) => {
			buffer[i] ^= b;
		});
		return digest('SHA-1', buffer);
	};
	const [x1, x2] = await Promise.all([expand(0x36), expand(0x5c)]);
	return concatBytes(x1, x2).slice(0, keyBits / 8);
}

/** The key is right when SHA-1 of the decrypted verifier matches the decrypted verifier hash. */
export async function verifyStandardKey(
	info: StandardEncryptionInfo,
	key: Uint8Array,
): Promise<boolean> {
	const hashBlocks = info.encryptedVerifierHash.subarray(0, 32);
	if (info.verifierHashSize !== SHA1_SIZE || hashBlocks.length !== 32) {
		throw damagedPackage('the password verifier has the wrong length');
	}
	const cipher = await aesCbc(key);
	const plain = await aesEcbDecrypt(cipher, concatBytes(info.encryptedVerifier, hashBlocks));
	const expected = await digest('SHA-1', plain.subarray(0, 16));
	return equalBytes(expected, plain.subarray(16, 16 + SHA1_SIZE));
}

/** Decrypt the package: AES-ECB over everything after the 8-byte size. */
export async function decryptStandardPackage(
	key: Uint8Array,
	encryptedPackage: Uint8Array,
): Promise<Uint8Array> {
	const size = encryptedPackageSize(encryptedPackage);
	const data = encryptedPackage.subarray(8);
	// Only whole blocks are encrypted data; a trailing partial block is slack.
	const blocks = data.subarray(0, data.length - (data.length % 16));
	if (blocks.length < size) {
		throw damagedPackage('the encrypted data is shorter than the package');
	}
	const out = await aesEcbDecrypt(await aesCbc(key), blocks);
	return out.subarray(0, size);
}
