/**
 * Password-encrypted OOXML packages, as Excel writes them with File > Info >
 * Protect Workbook > Encrypt with Password: an OLE2 compound file holding the
 * `EncryptionInfo` and `EncryptedPackage` streams and the `\u0006DataSpaces`
 * storage, around the encrypted ZIP package.
 *
 * Adapted from pptx-viewer (https://github.com/ChristopherVR/pptx-viewer), Apache-2.0.
 */
import { readCompoundFile, type CompoundFile } from '../cfb/read';
import { writeCompoundFile } from '../cfb/write';
import { isCompoundFile, toUint8Array, type BinaryInput } from '../opc-package';
import {
	agileSecretKey,
	checkAgileIntegrity,
	checkAgileSupport,
	decryptAgilePackage,
} from './agile';
import { encryptAgilePackage, type AgileEncryptOptions } from './agile-encrypt';
import { dataSpacesStreams } from './data-spaces';
import { parseEncryptionInfo, type PackageEncryptionInfo } from './encryption-info';
import { incorrectPassword, passwordRequired, unsupportedEncryption } from './errors';
import {
	decryptStandardPackage,
	standardKey,
	standardKeyBits,
	verifyStandardKey,
} from './standard';

/** Options of {@link encryptPackage}; the defaults match Excel 2016 and later. */
export type EncryptPackageOptions = AgileEncryptOptions;

function hasStream(cfb: CompoundFile, name: string): boolean {
	return cfb.entries.some((e) => e.type === 'stream' && e.path.toLowerCase() === name);
}

/** The bytes are a compound file with `EncryptionInfo` and `EncryptedPackage` streams. */
export function isEncryptedPackageFile(data: BinaryInput): boolean {
	const bytes = toUint8Array(data);
	if (!isCompoundFile(bytes)) {
		return false;
	}
	try {
		const cfb = readCompoundFile(bytes);
		return hasStream(cfb, 'encryptioninfo') && hasStream(cfb, 'encryptedpackage');
	} catch {
		return false;
	}
}

function openContainer(bytes: Uint8Array): { info: PackageEncryptionInfo; data: Uint8Array } {
	let cfb: CompoundFile;
	try {
		cfb = readCompoundFile(bytes);
	} catch (e) {
		const reason = e instanceof Error ? e.message : String(e);
		throw unsupportedEncryption(`the encrypted container cannot be read (${reason})`);
	}
	const info = cfb.stream('EncryptionInfo');
	const data = cfb.stream('EncryptedPackage');
	if (!info || !data) {
		throw unsupportedEncryption('the EncryptionInfo or EncryptedPackage stream is missing');
	}
	return { info: parseEncryptionInfo(info), data };
}

/**
 * How an encrypted package is encrypted (for diagnostics), or `undefined` when
 * the bytes are not one. Unsupported schemes throw an `XlsxPasswordError`.
 */
export function encryptionInfoOf(data: BinaryInput): PackageEncryptionInfo | undefined {
	const bytes = toUint8Array(data);
	return isEncryptedPackageFile(bytes) ? openContainer(bytes).info : undefined;
}

/**
 * Decrypt an encrypted package and return the ZIP bytes inside. Throws an
 * `XlsxPasswordError`: `unsupported` for schemes that cannot be decrypted here
 * (checked first), `required` without a password, `incorrect` for a wrong one.
 * A failed data integrity check throws an `XlsxFormatError`.
 */
export async function decryptPackage(
	cfbBytes: BinaryInput,
	password?: string,
): Promise<Uint8Array> {
	const { info, data } = openContainer(toUint8Array(cfbBytes));
	if (info.type === 'agile') {
		checkAgileSupport(info);
	} else {
		standardKeyBits(info);
	}
	if (password === undefined) {
		throw passwordRequired();
	}
	if (info.type === 'standard') {
		const key = await standardKey(info, password);
		if (!(await verifyStandardKey(info, key))) {
			throw incorrectPassword();
		}
		return decryptStandardPackage(key, data);
	}
	const secretKey = await agileSecretKey(info, password);
	if (!secretKey) {
		throw incorrectPassword();
	}
	await checkAgileIntegrity(info, secretKey, data);
	return decryptAgilePackage(info, secretKey, data);
}

/**
 * Encrypt ZIP package bytes with a password as Excel does (Agile encryption,
 * AES-256-CBC, SHA-512, 100000 spins, data integrity HMAC) and wrap them in a
 * compound file with the `\u0006DataSpaces` storage.
 */
export async function encryptPackage(
	zipBytes: BinaryInput,
	password: string,
	options: EncryptPackageOptions = {},
): Promise<Uint8Array> {
	const zip = toUint8Array(zipBytes);
	const { encryptionInfo, encryptedPackage } = await encryptAgilePackage(zip, password, options);
	return writeCompoundFile([
		...dataSpacesStreams(),
		{ path: 'EncryptionInfo', data: encryptionInfo },
		{ path: 'EncryptedPackage', data: encryptedPackage },
	]);
}
