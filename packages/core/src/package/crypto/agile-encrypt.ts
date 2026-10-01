/**
 * Writing ECMA-376 Agile encryption the way Excel 2016 and later do: AES-256-CBC,
 * SHA-512, 100000 spins, 4096-byte segments and the data integrity HMAC.
 *
 * Adapted from pptx-viewer (https://github.com/ChristopherVR/pptx-viewer), Apache-2.0.
 */
import {
	BLOCK_KEYS,
	SEGMENT_SIZE,
	blockIv,
	blockKeyCipher,
	checkCipherParams,
	spinPasswordHash,
} from './agile';
import {
	buildAgileEncryptionInfo,
	type AgileCipherParams,
	type AgileEncryptionInfo,
	type AgilePasswordKey,
} from './encryption-info';
import {
	HASH_SIZE,
	aesCbc,
	digest,
	fitBytes,
	hmac,
	padToBlock,
	randomBytes,
	uint32LE,
	type HashAlgorithm,
} from './primitives';

/** Options for writing agile encryption; the defaults are what Excel 2016 and later write. */
export interface AgileEncryptOptions {
	/** AES key size (default 256). */
	keyBits?: 128 | 192 | 256;
	/** Hash for key derivation and integrity (default SHA-512). */
	hashAlgorithm?: HashAlgorithm;
	/** Password hashing rounds (default 100000). */
	spinCount?: number;
}

/** Encrypt a ZIP package; returns the EncryptionInfo and EncryptedPackage streams. */
export async function encryptAgilePackage(
	zip: Uint8Array,
	password: string,
	options: AgileEncryptOptions = {},
): Promise<{ encryptionInfo: Uint8Array; encryptedPackage: Uint8Array }> {
	const hash = options.hashAlgorithm ?? 'SHA-512';
	const keyBits = options.keyBits ?? 256;
	const spinCount = options.spinCount ?? 100_000;
	const params = (saltValue: Uint8Array): AgileCipherParams => ({
		saltSize: saltValue.length,
		blockSize: 16,
		keyBits,
		hashSize: HASH_SIZE[hash],
		cipherAlgorithm: 'AES',
		cipherChaining: 'ChainingModeCBC',
		hashAlgorithm: hash.replace('-', ''),
		saltValue,
	});
	const keyData = params(randomBytes(16));
	const secretKey = randomBytes(keyBits / 8);
	const keySalt = params(randomBytes(16));
	const passwordKey = await passwordKeyEncryptor(keySalt, password, spinCount, secretKey);

	const cipher = await aesCbc(secretKey);
	const padded = padToBlock(zip, 16);
	const encryptedPackage = new Uint8Array(8 + padded.length);
	const view = new DataView(encryptedPackage.buffer);
	view.setUint32(0, zip.length % 0x100000000, true);
	view.setUint32(4, Math.floor(zip.length / 0x100000000), true);
	await Promise.all(
		Array.from({ length: Math.ceil(padded.length / SEGMENT_SIZE) }, async (_, s) => {
			const segment = padded.subarray(s * SEGMENT_SIZE, (s + 1) * SEGMENT_SIZE);
			const encrypted = await cipher.encrypt(await blockIv(keyData, hash, uint32LE(s)), segment);
			encryptedPackage.set(encrypted, 8 + s * SEGMENT_SIZE);
		}),
	);

	// Data integrity: an HMAC of the whole stream under a random key, both encrypted.
	const hmacKey = randomBytes(HASH_SIZE[hash]);
	const hmacValue = await hmac(hash, hmacKey, encryptedPackage);
	const seal = async (blockKey: Uint8Array, data: Uint8Array): Promise<Uint8Array> =>
		cipher.encrypt(await blockIv(keyData, hash, blockKey), padToBlock(data, 16));
	const dataIntegrity = {
		encryptedHmacKey: await seal(BLOCK_KEYS.hmacKey, hmacKey),
		encryptedHmacValue: await seal(BLOCK_KEYS.hmacValue, hmacValue),
	};
	const info: AgileEncryptionInfo = {
		type: 'agile',
		version: '4.4',
		keyData,
		dataIntegrity,
		passwordKey,
	};
	return { encryptionInfo: buildAgileEncryptionInfo(info), encryptedPackage };
}

/** The password key encryptor: verifier input, its hash and the secret key, encrypted. */
async function passwordKeyEncryptor(
	params: AgileCipherParams,
	password: string,
	spinCount: number,
	secretKey: Uint8Array,
): Promise<AgilePasswordKey> {
	const hash = checkCipherParams(params);
	const spun = spinPasswordHash(hash, params.saltValue, password, spinCount);
	const iv = fitBytes(params.saltValue, params.blockSize, 0x36);
	const encrypt = async (blockKey: Uint8Array, data: Uint8Array): Promise<Uint8Array> => {
		const cipher = await blockKeyCipher(hash, spun, blockKey, params.keyBits);
		return cipher.encrypt(iv, padToBlock(data, 16));
	};
	const verifierInput = randomBytes(params.saltSize);
	return {
		...params,
		spinCount,
		encryptedVerifierHashInput: await encrypt(BLOCK_KEYS.verifierHashInput, verifierInput),
		encryptedVerifierHashValue: await encrypt(
			BLOCK_KEYS.verifierHashValue,
			await digest(hash, verifierInput),
		),
		encryptedKeyValue: await encrypt(BLOCK_KEYS.encryptedKeyValue, secretKey),
	};
}
