/**
 * Password-encrypted (ECMA-376) packages: decrypt to the ZIP inside, encrypt a
 * ZIP as Excel does, and inspect how a file is encrypted.
 */
export {
	decryptPackage,
	encryptPackage,
	encryptionInfoOf,
	isEncryptedPackageFile,
	type EncryptPackageOptions,
} from './package-crypto';
export type {
	AgileCipherParams,
	AgileEncryptionInfo,
	AgilePasswordKey,
	PackageEncryptionInfo,
	StandardEncryptionInfo,
} from './encryption-info';
export type { HashAlgorithm } from './sha';
