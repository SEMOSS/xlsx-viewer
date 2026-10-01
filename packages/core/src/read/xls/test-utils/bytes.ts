/** Little-endian byte encoders for the test-only BIFF8 builders. */

export function u16(v: number): number[] {
	return [v & 0xff, (v >>> 8) & 0xff];
}

export function u32(v: number): number[] {
	return [v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff];
}

export function f64(v: number): number[] {
	const b = new Uint8Array(8);
	new DataView(b.buffer).setFloat64(0, v, true);
	return [...b];
}

/** Characters as 8-bit (compressed) or UTF-16LE code units. */
export function chars(s: string, high: boolean): number[] {
	const out: number[] = [];
	for (let i = 0; i < s.length; i++) {
		const c = s.charCodeAt(i);
		if (high) {
			out.push(c & 0xff, c >>> 8);
		} else {
			out.push(c & 0xff);
		}
	}
	return out;
}

/** Strings with characters above U+00FF must be stored as 16-bit. */
export function needsHigh(s: string): boolean {
	return [...s].some((c) => c.charCodeAt(0) > 0xff);
}

/** `XLUnicodeString`: 16-bit count, option byte, characters. */
export function xlString(s: string, high = needsHigh(s)): number[] {
	return [...u16(s.length), high ? 1 : 0, ...chars(s, high)];
}

/** `ShortXLUnicodeString`: 8-bit count, option byte, characters. */
export function shortString(s: string, high = needsHigh(s)): number[] {
	return [s.length, high ? 1 : 0, ...chars(s, high)];
}

/** `XLUnicodeStringNoCch`: option byte and characters. */
export function stringNoCch(s: string, high = needsHigh(s)): number[] {
	return [high ? 1 : 0, ...chars(s, high)];
}

/** A BIFF record: type, size and data (no automatic CONTINUE splitting). */
export function record(type: number, data: readonly number[]): number[] {
	return [...u16(type), ...u16(data.length), ...data];
}

/** Concatenate byte arrays. */
export function concat(parts: ReadonlyArray<readonly number[] | Uint8Array>): Uint8Array {
	const total = parts.reduce((n, p) => n + p.length, 0);
	const out = new Uint8Array(total);
	let at = 0;
	for (const p of parts) {
		out.set(p, at);
		at += p.length;
	}
	return out;
}

/** `Ref8`: first row, last row, first column, last column (16 bits each). */
export const ref8 = (r1: number, r2: number, c1: number, c2: number): number[] => [
	...u16(r1),
	...u16(r2),
	...u16(c1),
	...u16(c2),
];
