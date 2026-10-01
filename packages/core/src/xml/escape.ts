const NAMED_ENTITIES: Record<string, string> = {
	lt: '<',
	gt: '>',
	amp: '&',
	quot: '"',
	apos: "'",
};

/** Decode XML character and entity references. */
export function decodeXmlEntities(text: string): string {
	if (text.indexOf('&') === -1) {
		return text;
	}
	return text.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]+);/g, (whole, body: string) => {
		if (body.charCodeAt(0) === 35 /* # */) {
			const code =
				body.charCodeAt(1) === 120 /* x */ || body.charCodeAt(1) === 88
					? parseInt(body.slice(2), 16)
					: parseInt(body.slice(1), 10);
			if (!Number.isFinite(code) || code < 0 || code > 0x10ffff) {
				return whole;
			}
			return String.fromCodePoint(code);
		}
		const named = NAMED_ENTITIES[body];
		return named ?? whole;
	});
}

/** Escape text content. */
export function escapeXmlText(text: string): string {
	let out = '';
	let last = 0;
	for (let i = 0; i < text.length; i++) {
		const c = text.charCodeAt(i);
		let rep: string | null = null;
		if (c === 38) {
			rep = '&amp;';
		} else if (c === 60) {
			rep = '&lt;';
		} else if (c === 62) {
			rep = '&gt;';
		} else if (c < 32 && c !== 9 && c !== 10 && c !== 13) {
			// Control characters are not allowed in XML 1.0; Excel writes them as _x00HH_.
			rep = `_x${c.toString(16).padStart(4, '0').toUpperCase()}_`;
		}
		if (rep !== null) {
			out += text.slice(last, i) + rep;
			last = i + 1;
		}
	}
	return last === 0 ? text : out + text.slice(last);
}

/** Escape an attribute value (double-quoted). Newlines and tabs are encoded to survive normalisation. */
export function escapeXmlAttr(text: string): string {
	let out = '';
	let last = 0;
	for (let i = 0; i < text.length; i++) {
		const c = text.charCodeAt(i);
		let rep: string | null = null;
		switch (c) {
			case 38:
				rep = '&amp;';
				break;
			case 60:
				rep = '&lt;';
				break;
			case 62:
				rep = '&gt;';
				break;
			case 34:
				rep = '&quot;';
				break;
			case 10:
				rep = '&#10;';
				break;
			case 13:
				rep = '&#13;';
				break;
			case 9:
				rep = '&#9;';
				break;
			default:
				if (c < 32) {
					rep = `_x${c.toString(16).padStart(4, '0').toUpperCase()}_`;
				}
		}
		if (rep !== null) {
			out += text.slice(last, i) + rep;
			last = i + 1;
		}
	}
	return last === 0 ? text : out + text.slice(last);
}

/**
 * Decode the OOXML `_xHHHH_` escape used for characters that XML cannot carry
 * (control characters) in shared strings and cell text. A literal `_x` sequence
 * is itself escaped as `_x005F_x...`.
 */
export function decodeOoxmlEscapes(text: string): string {
	if (text.indexOf('_x') === -1) {
		return text;
	}
	return text.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, hex: string) =>
		String.fromCharCode(parseInt(hex, 16)),
	);
}

/** Encode literal `_xHHHH_` sequences so they survive {@link decodeOoxmlEscapes}. */
export function encodeOoxmlEscapes(text: string): string {
	if (text.indexOf('_x') === -1) {
		return text;
	}
	return text.replace(/_x([0-9A-Fa-f]{4})_/g, '_x005F_x$1_');
}
