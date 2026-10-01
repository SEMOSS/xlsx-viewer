/**
 * The `EncryptionInfo` stream ([MS-OFFCRYPTO] 2.3.4.5 and 2.3.4.10): a
 * version header, then either the binary Standard encryption header and
 * verifier, or the XML descriptor of Agile encryption.
 *
 * Adapted from pptx-viewer (https://github.com/ChristopherVR/pptx-viewer), Apache-2.0.
 */
import {
	attr,
	childLocal,
	childrenLocal,
	localName,
	parseXmlRoot,
	type XmlElement,
} from '../../xml';
import { damagedPackage, unsupportedEncryption } from './errors';
import { base64Decode, base64Encode } from './primitives';

const ENCRYPTION_NS = 'http://schemas.microsoft.com/office/2006/encryption';
export const PASSWORD_KEY_ENCRYPTOR =
	'http://schemas.microsoft.com/office/2006/keyEncryptor/password';
const CERTIFICATE_KEY_ENCRYPTOR =
	'http://schemas.microsoft.com/office/2006/keyEncryptor/certificate';

/** Cipher parameters of the agile `keyData` and password key encryptor elements. */
export interface AgileCipherParams {
	saltSize: number;
	blockSize: number;
	keyBits: number;
	hashSize: number;
	/** `AES` (the only cipher supported here). */
	cipherAlgorithm: string;
	/** `ChainingModeCBC` (supported) or `ChainingModeCFB`. */
	cipherChaining: string;
	/** As written in the file: `SHA512`, `SHA1`, `SHA-1`, ... */
	hashAlgorithm: string;
	saltValue: Uint8Array;
}

/** The password `p:encryptedKey` element. */
export interface AgilePasswordKey extends AgileCipherParams {
	spinCount: number;
	encryptedVerifierHashInput: Uint8Array;
	encryptedVerifierHashValue: Uint8Array;
	encryptedKeyValue: Uint8Array;
}

/** Agile encryption (Excel 2010 and later), version 4.4. */
export interface AgileEncryptionInfo {
	type: 'agile';
	version: string;
	keyData: AgileCipherParams;
	dataIntegrity?: { encryptedHmacKey: Uint8Array; encryptedHmacValue: Uint8Array };
	/** Absent when the file can only be opened with a certificate. */
	passwordKey?: AgilePasswordKey;
}

/** Standard encryption (Excel 2007): versions 2.2, 3.2 and 4.2. */
export interface StandardEncryptionInfo {
	type: 'standard';
	version: string;
	/** EncryptionHeader flags: `fCryptoAPI` 0x04, `fExternal` 0x10, `fAES` 0x20. */
	flags: number;
	/** CryptoAPI algorithm: 0x660E / 0x660F / 0x6610 for AES-128 / 192 / 256, 0x6801 for RC4. */
	algId: number;
	/** 0x8004 (SHA-1), or 0 for the default. */
	algIdHash: number;
	keyBits: number;
	providerType: number;
	cspName: string;
	salt: Uint8Array;
	encryptedVerifier: Uint8Array;
	verifierHashSize: number;
	encryptedVerifierHash: Uint8Array;
}

export type PackageEncryptionInfo = AgileEncryptionInfo | StandardEncryptionInfo;

/** Parse an `EncryptionInfo` stream; unsupported kinds throw an `XlsxPasswordError`. */
export function parseEncryptionInfo(stream: Uint8Array): PackageEncryptionInfo {
	if (stream.length < 8) {
		throw unsupportedEncryption('the EncryptionInfo stream is truncated');
	}
	const view = new DataView(stream.buffer, stream.byteOffset, stream.byteLength);
	const major = view.getUint16(0, true);
	const minor = view.getUint16(2, true);
	const version = `${major}.${minor}`;
	if (major === 4 && minor === 4) {
		return parseAgile(stream.subarray(8), version);
	}
	if ((major === 3 || major === 4) && minor === 3) {
		throw unsupportedEncryption('extensible encryption (a third-party encryption module)');
	}
	if (major >= 2 && major <= 4 && minor === 2) {
		return parseStandard(stream, view, version);
	}
	throw unsupportedEncryption(`unknown encryption version ${version}`);
}

function parseStandard(
	stream: Uint8Array,
	view: DataView,
	version: string,
): StandardEncryptionInfo {
	const headerSize = view.getUint32(8, true);
	const verifier = 12 + headerSize;
	if (headerSize < 32 || verifier + 40 > stream.length) {
		throw damagedPackage('the encryption header is truncated');
	}
	if ((view.getUint32(4, true) | view.getUint32(12, true)) & 0x10) {
		throw unsupportedEncryption('extensible encryption (a third-party encryption module)');
	}
	let cspName = '';
	for (let i = 12 + 32; i + 1 < verifier; i += 2) {
		const c = view.getUint16(i, true);
		if (c === 0) {
			break;
		}
		cspName += String.fromCharCode(c);
	}
	// EncryptionVerifier: salt size (16), salt, encrypted verifier, hash size, encrypted hash.
	const saltSize = view.getUint32(verifier, true);
	if (saltSize !== 16) {
		throw damagedPackage(`the encryption verifier has a ${saltSize}-byte salt`);
	}
	const verifierHashSize = view.getUint32(verifier + 36, true);
	const hashStart = verifier + 40;
	return {
		type: 'standard',
		version,
		flags: view.getUint32(12, true),
		algId: view.getUint32(20, true),
		algIdHash: view.getUint32(24, true),
		keyBits: view.getUint32(28, true),
		providerType: view.getUint32(32, true),
		cspName,
		salt: stream.slice(verifier + 4, verifier + 20),
		encryptedVerifier: stream.slice(verifier + 20, verifier + 36),
		verifierHashSize,
		// AES pads the 20-byte SHA-1 verifier hash to 32 bytes; RC4 stores it as is.
		encryptedVerifierHash: stream.slice(hashStart),
	};
}

function numberAttr(el: XmlElement, name: string): number {
	const value = Number(attr(el, name));
	if (!Number.isInteger(value) || value < 0) {
		throw damagedPackage(`the encryption descriptor has no valid ${name}`);
	}
	return value;
}

function bytesAttr(el: XmlElement, name: string): Uint8Array {
	try {
		return base64Decode(attr(el, name) ?? '');
	} catch {
		throw damagedPackage(`the encryption descriptor has an invalid ${name}`);
	}
}

function cipherParams(el: XmlElement): AgileCipherParams {
	return {
		saltSize: numberAttr(el, 'saltSize'),
		blockSize: numberAttr(el, 'blockSize'),
		keyBits: numberAttr(el, 'keyBits'),
		hashSize: numberAttr(el, 'hashSize'),
		cipherAlgorithm: attr(el, 'cipherAlgorithm') ?? '',
		cipherChaining: attr(el, 'cipherChaining') ?? '',
		hashAlgorithm: attr(el, 'hashAlgorithm') ?? '',
		saltValue: bytesAttr(el, 'saltValue'),
	};
}

function parseAgile(xmlBytes: Uint8Array, version: string): AgileEncryptionInfo {
	let root: XmlElement;
	try {
		// Some writers pad the stream after the XML.
		const text = new TextDecoder('utf-8').decode(xmlBytes).replace(/[\0\s]+$/, '');
		root = parseXmlRoot(text);
	} catch {
		throw damagedPackage('the encryption descriptor is not valid XML');
	}
	const keyDataEl = childLocal(root, 'keyData');
	if (localName(root.name) !== 'encryption' || !keyDataEl) {
		throw damagedPackage('the encryption descriptor has no key data');
	}
	const info: AgileEncryptionInfo = { type: 'agile', version, keyData: cipherParams(keyDataEl) };
	const integrity = childLocal(root, 'dataIntegrity');
	if (integrity) {
		info.dataIntegrity = {
			encryptedHmacKey: bytesAttr(integrity, 'encryptedHmacKey'),
			encryptedHmacValue: bytesAttr(integrity, 'encryptedHmacValue'),
		};
	}
	const encryptors = childrenLocal(childLocal(root, 'keyEncryptors'), 'keyEncryptor');
	const password = encryptors.find((e) => attr(e, 'uri') === PASSWORD_KEY_ENCRYPTOR);
	const key = childLocal(password, 'encryptedKey');
	if (key) {
		info.passwordKey = {
			...cipherParams(key),
			spinCount: numberAttr(key, 'spinCount'),
			encryptedVerifierHashInput: bytesAttr(key, 'encryptedVerifierHashInput'),
			encryptedVerifierHashValue: bytesAttr(key, 'encryptedVerifierHashValue'),
			encryptedKeyValue: bytesAttr(key, 'encryptedKeyValue'),
		};
	}
	return info;
}

function paramsXml(p: AgileCipherParams): string {
	return (
		`saltSize="${p.saltSize}" blockSize="${p.blockSize}" keyBits="${p.keyBits}" ` +
		`hashSize="${p.hashSize}" cipherAlgorithm="${p.cipherAlgorithm}" ` +
		`cipherChaining="${p.cipherChaining}" hashAlgorithm="${p.hashAlgorithm}" ` +
		`saltValue="${base64Encode(p.saltValue)}"`
	);
}

/**
 * Serialize agile encryption info the way Excel does: version 4.4, reserved
 * 0x40, then the UTF-8 XML descriptor.
 */
export function buildAgileEncryptionInfo(info: AgileEncryptionInfo): Uint8Array {
	const key = info.passwordKey;
	const integrity = info.dataIntegrity;
	if (!key || !integrity) {
		throw new Error('Agile encryption info needs a password key and data integrity values');
	}
	const xml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
		`<encryption xmlns="${ENCRYPTION_NS}" xmlns:p="${PASSWORD_KEY_ENCRYPTOR}" ` +
		`xmlns:c="${CERTIFICATE_KEY_ENCRYPTOR}">` +
		`<keyData ${paramsXml(info.keyData)}/>` +
		`<dataIntegrity encryptedHmacKey="${base64Encode(integrity.encryptedHmacKey)}" ` +
		`encryptedHmacValue="${base64Encode(integrity.encryptedHmacValue)}"/>` +
		`<keyEncryptors><keyEncryptor uri="${PASSWORD_KEY_ENCRYPTOR}">` +
		`<p:encryptedKey spinCount="${key.spinCount}" ${paramsXml(key)} ` +
		`encryptedVerifierHashInput="${base64Encode(key.encryptedVerifierHashInput)}" ` +
		`encryptedVerifierHashValue="${base64Encode(key.encryptedVerifierHashValue)}" ` +
		`encryptedKeyValue="${base64Encode(key.encryptedKeyValue)}"/>` +
		'</keyEncryptor></keyEncryptors></encryption>';
	const body = new TextEncoder().encode(xml);
	const out = new Uint8Array(8 + body.length);
	out.set([4, 0, 4, 0, 0x40, 0, 0, 0]);
	out.set(body, 8);
	return out;
}
