/**
 * Errors of the encryption layer, as the workbook reader reports them:
 * password problems are `XlsxPasswordError`s (`required`, `incorrect`,
 * `unsupported`), damaged containers are `XlsxFormatError`s.
 */
import { XlsxFormatError, XlsxPasswordError } from '../../read/errors';

export function passwordRequired(): XlsxPasswordError {
	return new XlsxPasswordError('This workbook is protected with a password.', 'required');
}

export function incorrectPassword(): XlsxPasswordError {
	return new XlsxPasswordError('The password you supplied is not correct.', 'incorrect');
}

export function unsupportedEncryption(detail: string): XlsxPasswordError {
	return new XlsxPasswordError(
		`This workbook is encrypted in a way that is not supported: ${detail}.`,
		'unsupported',
	);
}

export function damagedPackage(detail: string): XlsxFormatError {
	return new XlsxFormatError(`The encrypted workbook is damaged: ${detail}.`);
}

export function integrityCheckFailed(): XlsxFormatError {
	return new XlsxFormatError(
		'The encrypted workbook failed its data integrity check (HMAC mismatch): ' +
			'the file is damaged or was changed after it was saved.',
	);
}
