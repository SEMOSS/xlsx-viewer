import { quoteSheetName } from '../address/sheet-name';

/**
 * Hyperlink address text as the Insert Hyperlink dialog handles it: web
 * and e-mail addresses typed without a scheme, mailto: links with a
 * subject, and "Place in This Document" locations (`'My Sheet'!A1`).
 */

/** A location inside the workbook: an optional sheet and a cell reference or defined name. */
export interface HyperlinkLocation {
	sheet?: string;
	ref: string;
}

const SCHEME = /^[a-z][a-z0-9+.-]*:/i;
const EMAIL = /^[^\s@/\\:?#]+@[^\s@/\\:?#]+\.[^\s@/\\:?#]+$/;
const MAILTO = /^mailto:/i;

/**
 * Complete a typed address like Excel: `www.example.com` becomes
 * `http://www.example.com`, `ftp.example.com` becomes
 * `ftp://ftp.example.com` and `user@example.com` becomes
 * `mailto:user@example.com`. Addresses with a scheme (`https:`,
 * `mailto:`, `file:`), drive and UNC paths and relative file names stay.
 */
export function normalizeHyperlinkAddress(address: string): string {
	const text = address.trim();
	if (text === '' || /^[\\/.#]/.test(text)) {
		return text;
	}
	if (/^www\./i.test(text)) {
		return `http://${text}`;
	}
	if (/^ftp\./i.test(text)) {
		return `ftp://${text}`;
	}
	if (SCHEME.test(text)) {
		return text;
	}
	return EMAIL.test(text) ? `mailto:${text}` : text;
}

/** `mailto:` link for the E-mail Address page: `mailto:a@b.com?subject=Hello%20there`. */
export function buildMailtoAddress(email: string, subject?: string): string {
	const address = email.trim().replace(MAILTO, '');
	const base = `mailto:${address}`;
	return subject && subject.trim() !== '' ? `${base}?subject=${encodeURIComponent(subject)}` : base;
}

function decode(text: string): string {
	try {
		return decodeURIComponent(text.replace(/\+/g, '%20'));
	} catch {
		return text;
	}
}

/** Split a `mailto:` target into address and subject; null for other targets. */
export function parseMailtoAddress(target: string): { email: string; subject: string } | null {
	if (!MAILTO.test(target)) {
		return null;
	}
	const rest = target.replace(MAILTO, '');
	const q = rest.indexOf('?');
	const email = decode(q < 0 ? rest : rest.slice(0, q));
	let subject = '';
	if (q >= 0) {
		for (const pair of rest.slice(q + 1).split('&')) {
			const eq = pair.indexOf('=');
			if (eq > 0 && pair.slice(0, eq).toLowerCase() === 'subject') {
				subject = decode(pair.slice(eq + 1));
			}
		}
	}
	return { email, subject };
}

/**
 * Split a hyperlink location (`Sheet2!B4`, `'My Sheet'!A1:C3`, `Rate`,
 * `#Sheet1!A1`) into its sheet and the reference or name after it.
 */
export function parseHyperlinkLocation(location: string): HyperlinkLocation {
	let text = location.trim();
	if (text.startsWith('#')) {
		text = text.slice(1);
	}
	if (text.startsWith("'")) {
		let i = 1;
		while (i < text.length) {
			if (text[i] === "'") {
				if (text[i + 1] !== "'") {
					break;
				}
				i++;
			}
			i++;
		}
		if (text[i] === "'" && text[i + 1] === '!') {
			return { sheet: text.slice(1, i).replace(/''/g, "'"), ref: text.slice(i + 2) };
		}
		return { ref: text };
	}
	const bang = text.lastIndexOf('!');
	return bang > 0 ? { sheet: text.slice(0, bang), ref: text.slice(bang + 1) } : { ref: text };
}

/** Location text for a sheet and reference: `Sheet2!A1`, `'My Sheet'!A1`, or `ref` alone. */
export function formatHyperlinkLocation(sheetName: string | undefined, ref: string): string {
	return sheetName ? `${quoteSheetName(sheetName)}!${ref}` : ref;
}
