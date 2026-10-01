import { compoundFileEntryNames } from './cfb';

/** Thrown when a file is password-protected and no (or a wrong) password was given. */
export class XlsxPasswordError extends Error {
	constructor(
		message: string,
		readonly reason: 'required' | 'incorrect' | 'unsupported',
	) {
		super(message);
		this.name = 'XlsxPasswordError';
	}
}

/** Thrown for inputs that are not a readable workbook (legacy `.xls`, corrupt ZIP, ...). */
export class XlsxFormatError extends Error {
	constructor(message: string) {
		super(message);
		this.name = 'XlsxFormatError';
	}
}

/** UTF-16LE bytes of a compound-file directory entry name. */
function utf16le(name: string): Uint8Array {
	const out = new Uint8Array(name.length * 2);
	for (let i = 0; i < name.length; i++) {
		out[i * 2] = name.charCodeAt(i) & 0xff;
		out[i * 2 + 1] = name.charCodeAt(i) >> 8;
	}
	return out;
}

const ENCRYPTION_INFO = utf16le('EncryptionInfo');
const ENCRYPTED_PACKAGE = utf16le('EncryptedPackage');

function containsBytes(haystack: Uint8Array, needle: Uint8Array): boolean {
	const first = needle[0];
	const last = haystack.length - needle.length;
	outer: for (let i = 0; i <= last; i++) {
		if (haystack[i] !== first) {
			continue;
		}
		for (let j = 1; j < needle.length; j++) {
			if (haystack[i + j] !== needle[j]) {
				continue outer;
			}
		}
		return true;
	}
	return false;
}

/**
 * The compound file holds an ECMA-376 encrypted package (Encrypt with Password),
 * rather than a legacy workbook.
 */
export function isEncryptedPackage(bytes: Uint8Array): boolean {
	const names = compoundFileEntryNames(bytes);
	if (names) {
		return names.includes('EncryptionInfo') || names.includes('EncryptedPackage');
	}
	// Unreadable directory: fall back to looking for the UTF-16LE stream names anywhere.
	return containsBytes(bytes, ENCRYPTION_INFO) || containsBytes(bytes, ENCRYPTED_PACKAGE);
}

/**
 * The error to throw for an OLE2 compound file the package reader cannot
 * open: an encrypted OOXML package without its password (ECMA-376 encryption
 * stores `EncryptionInfo` / `EncryptedPackage` streams), or a compound file
 * that is not a workbook. Legacy `.xls` files go to the `.xls` reader first.
 */
export function compoundFileError(bytes: Uint8Array): Error {
	if (isEncryptedPackage(bytes)) {
		// readWorkbook decrypts these; without a password this is the error it reports.
		return new XlsxPasswordError('This workbook is protected with a password.', 'required');
	}
	return new XlsxFormatError('The file is an OLE2 compound document, not an Excel workbook');
}
