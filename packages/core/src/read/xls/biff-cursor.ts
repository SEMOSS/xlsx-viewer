/**
 * Bounds-checked little-endian reader over the data of one BIFF record,
 * including the data of the CONTINUE records that followed it. `breaks`
 * are the offsets where each CONTINUE record's data starts: a string whose
 * characters cross a break carries a fresh option byte there (its
 * `fHighByte` flag selects 8-bit "compressed" or 16-bit characters for the
 * rest of the string). Every other value reads straight across breaks.
 */

/** A read went past the end of the record data. */
export class BiffTruncatedError extends Error {
	constructor(what = 'record') {
		super(`The ${what} is truncated`);
		this.name = 'BiffTruncatedError';
	}
}

const CHUNK = 4096;

/** Decode `count` characters stored as 8-bit (Latin-1) or UTF-16LE code units. */
export function decodeChars(data: Uint8Array, start: number, count: number, high: boolean): string {
	let out = '';
	const codes: number[] = [];
	for (let i = 0; i < count; i++) {
		const p = high ? start + i * 2 : start + i;
		codes.push(high ? data[p] | (data[p + 1] << 8) : data[p]);
		if (codes.length === CHUNK) {
			out += String.fromCharCode(...codes);
			codes.length = 0;
		}
	}
	return codes.length ? out + String.fromCharCode(...codes) : out;
}

export class BiffCursor {
	pos: number;
	private readonly view: DataView;
	/** Index of the first break at or after `pos` (kept in sync lazily). */
	private bi = 0;

	constructor(
		readonly data: Uint8Array,
		readonly breaks: readonly number[] = [],
		start = 0,
	) {
		this.view = new DataView(data.buffer, data.byteOffset, data.byteLength);
		this.pos = start;
	}

	get remaining(): number {
		return this.data.length - this.pos;
	}

	private take(n: number): number {
		const p = this.pos;
		if (n < 0 || p + n > this.data.length) {
			throw new BiffTruncatedError();
		}
		this.pos = p + n;
		return p;
	}

	u8(): number {
		return this.data[this.take(1)];
	}

	u16(): number {
		return this.view.getUint16(this.take(2), true);
	}

	i16(): number {
		return this.view.getInt16(this.take(2), true);
	}

	u32(): number {
		return this.view.getUint32(this.take(4), true);
	}

	i32(): number {
		return this.view.getInt32(this.take(4), true);
	}

	f64(): number {
		return this.view.getFloat64(this.take(8), true);
	}

	skip(n: number): void {
		this.take(n);
	}

	bytes(n: number): Uint8Array {
		const p = this.take(n);
		return this.data.subarray(p, p + n);
	}

	private syncBreaks(): void {
		const b = this.breaks;
		while (this.bi < b.length && b[this.bi] < this.pos) {
			this.bi++;
		}
		while (this.bi > 0 && b[this.bi - 1] >= this.pos) {
			this.bi--;
		}
	}

	/** Offset of the first CONTINUE break at or after `pos`, or the end of the data. */
	nextBreak(): number {
		this.syncBreaks();
		return this.bi < this.breaks.length ? this.breaks[this.bi] : this.data.length;
	}

	/**
	 * `count` characters. At a CONTINUE break inside the characters a new
	 * option byte chooses the character width for the remainder.
	 */
	chars(count: number, highByte: boolean): string {
		let out = '';
		let high = highByte;
		let left = count;
		while (left > 0) {
			this.syncBreaks();
			if (this.bi < this.breaks.length && this.breaks[this.bi] === this.pos) {
				high = (this.u8() & 1) !== 0;
				this.syncBreaks();
			}
			const limit = this.bi < this.breaks.length ? this.breaks[this.bi] : this.data.length;
			const width = high ? 2 : 1;
			const avail = Math.floor((limit - this.pos) / width);
			if (avail <= 0) {
				if (limit <= this.pos || limit === this.data.length) {
					throw new BiffTruncatedError('string');
				}
				// A stray byte before the break (malformed split): skip to the break.
				this.pos = limit;
				continue;
			}
			const n = Math.min(left, avail);
			out += decodeChars(this.data, this.pos, n, high);
			this.pos += n * width;
			left -= n;
		}
		return out;
	}

	/** `XLUnicodeString`: 16-bit character count, option byte, characters. */
	unicodeString(): string {
		const cch = this.u16();
		return this.chars(cch, (this.u8() & 1) !== 0);
	}

	/** `ShortXLUnicodeString`: 8-bit character count, option byte, characters. */
	shortString(): string {
		const cch = this.u8();
		return this.chars(cch, (this.u8() & 1) !== 0);
	}

	/** `XLUnicodeStringNoCch`: option byte and `count` characters (the count is stored elsewhere). */
	stringNoCch(count: number): string {
		return this.chars(count, (this.u8() & 1) !== 0);
	}

	/** A null-terminated UTF-16LE string (as in HLINKTOOLTIP), up to the end of the data. */
	utf16z(): string {
		let end = this.pos;
		while (end + 1 < this.data.length && (this.data[end] | this.data[end + 1]) !== 0) {
			end += 2;
		}
		const text = decodeChars(this.data, this.pos, (end - this.pos) >> 1, true);
		this.pos = Math.min(this.data.length, end + 2);
		return text;
	}
}
