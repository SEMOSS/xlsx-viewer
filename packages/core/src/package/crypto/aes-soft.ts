/**
 * AES (FIPS 197) block cipher in plain TypeScript, with CBC chaining and no
 * padding. A fallback for what WebCrypto cannot do everywhere: Chromium-based
 * browsers reject 192-bit AES keys, and insecure (non-HTTPS) pages have no
 * `crypto.subtle` at all. Uses the usual 32-bit lookup tables.
 */

const SBOX = new Uint8Array(256);
const INV_SBOX = new Uint8Array(256);
const TE = [0, 1, 2, 3].map(() => new Int32Array(256));
const TD = [0, 1, 2, 3].map(() => new Int32Array(256));

const xtime = (x: number): number => ((x << 1) ^ (x & 0x80 ? 0x1b : 0)) & 0xff;

let tablesReady = false;

/** Build the S-boxes from GF(2^8) inverses and the cipher tables from them (once). */
function initTables(): void {
	if (tablesReady) {
		return;
	}
	tablesReady = true;
	const exp = new Uint8Array(256);
	const log = new Uint8Array(256);
	for (let i = 0, x = 1; i < 255; i++) {
		exp[i] = x;
		log[x] = i;
		x ^= xtime(x); // multiply by the generator 3
	}
	for (let i = 0; i < 256; i++) {
		const inv = i === 0 ? 0 : exp[(255 - log[i]) % 255];
		let s = inv;
		for (let r = 1; r < 5; r++) {
			s ^= ((inv << r) | (inv >>> (8 - r))) & 0xff;
		}
		s ^= 0x63;
		SBOX[i] = s;
		INV_SBOX[s] = i;
	}
	const mul = (a: number, b: number): number => (a && b ? exp[(log[a] + log[b]) % 255] : 0);
	for (let i = 0; i < 256; i++) {
		const s = SBOX[i];
		const e = (mul(s, 2) << 24) | (s << 16) | (s << 8) | mul(s, 3);
		const v = INV_SBOX[i];
		const d = (mul(v, 14) << 24) | (mul(v, 9) << 16) | (mul(v, 13) << 8) | mul(v, 11);
		for (let t = 0; t < 4; t++) {
			TE[t][i] = (e >>> (8 * t)) | (e << (32 - 8 * t));
			TD[t][i] = (d >>> (8 * t)) | (d << (32 - 8 * t));
		}
	}
}

const subWord = (w: number): number =>
	(SBOX[w >>> 24] << 24) |
	(SBOX[(w >>> 16) & 0xff] << 16) |
	(SBOX[(w >>> 8) & 0xff] << 8) |
	SBOX[w & 0xff];

/** Round keys for encryption, and for decryption with the equivalent inverse cipher. */
function expandKey(key: Uint8Array): { rounds: number; enc: Int32Array; dec: Int32Array } {
	const nk = key.length / 4;
	if (nk !== 4 && nk !== 6 && nk !== 8) {
		throw new Error(`Invalid AES key length: ${key.length * 8} bits`);
	}
	const rounds = nk + 6;
	const enc = new Int32Array(4 * (rounds + 1));
	for (let i = 0; i < nk; i++) {
		enc[i] = (key[i * 4] << 24) | (key[i * 4 + 1] << 16) | (key[i * 4 + 2] << 8) | key[i * 4 + 3];
	}
	for (let i = nk, rcon = 1; i < enc.length; i++) {
		let t = enc[i - 1];
		if (i % nk === 0) {
			t = subWord((t << 8) | (t >>> 24)) ^ (rcon << 24);
			rcon = xtime(rcon);
		} else if (nk > 6 && i % nk === 4) {
			t = subWord(t);
		}
		enc[i] = enc[i - nk] ^ t;
	}
	const dec = new Int32Array(enc.length);
	for (let r = 0; r <= rounds; r++) {
		for (let c = 0; c < 4; c++) {
			const w = enc[(rounds - r) * 4 + c];
			dec[r * 4 + c] =
				r === 0 || r === rounds
					? w
					: TD[0][SBOX[w >>> 24]] ^
						TD[1][SBOX[(w >>> 16) & 0xff]] ^
						TD[2][SBOX[(w >>> 8) & 0xff]] ^
						TD[3][SBOX[w & 0xff]];
		}
	}
	return { rounds, enc, dec };
}

const [TE0, TE1, TE2, TE3] = TE;
const [TD0, TD1, TD2, TD3] = TD;

/** A full round for one output column: the bytes come from columns a, b, c, d. */
const mixE = (a: number, b: number, c: number, d: number, key: number): number =>
	TE0[a >>> 24] ^ TE1[(b >>> 16) & 0xff] ^ TE2[(c >>> 8) & 0xff] ^ TE3[d & 0xff] ^ key;
const mixD = (a: number, b: number, c: number, d: number, key: number): number =>
	TD0[a >>> 24] ^ TD1[(b >>> 16) & 0xff] ^ TD2[(c >>> 8) & 0xff] ^ TD3[d & 0xff] ^ key;

/** The final round (no column mixing) for one output column. */
function lastRound(
	box: Uint8Array,
	a: number,
	b: number,
	c: number,
	d: number,
	key: number,
): number {
	const bytes =
		(box[a >>> 24] << 24) |
		(box[(b >>> 16) & 0xff] << 16) |
		(box[(c >>> 8) & 0xff] << 8) |
		box[d & 0xff];
	return bytes ^ key;
}

/** Encrypt one block of four big-endian column words in place. */
function encryptBlock(st: Int32Array, k: Int32Array, rounds: number): void {
	let s0 = st[0] ^ k[0];
	let s1 = st[1] ^ k[1];
	let s2 = st[2] ^ k[2];
	let s3 = st[3] ^ k[3];
	let i = 4;
	for (let r = 1; r < rounds; r++, i += 4) {
		const t0 = mixE(s0, s1, s2, s3, k[i]);
		const t1 = mixE(s1, s2, s3, s0, k[i + 1]);
		const t2 = mixE(s2, s3, s0, s1, k[i + 2]);
		const t3 = mixE(s3, s0, s1, s2, k[i + 3]);
		s0 = t0;
		s1 = t1;
		s2 = t2;
		s3 = t3;
	}
	st[0] = lastRound(SBOX, s0, s1, s2, s3, k[i]);
	st[1] = lastRound(SBOX, s1, s2, s3, s0, k[i + 1]);
	st[2] = lastRound(SBOX, s2, s3, s0, s1, k[i + 2]);
	st[3] = lastRound(SBOX, s3, s0, s1, s2, k[i + 3]);
}

/** Decrypt one block in place (equivalent inverse cipher: rows shift the other way). */
function decryptBlock(st: Int32Array, k: Int32Array, rounds: number): void {
	let s0 = st[0] ^ k[0];
	let s1 = st[1] ^ k[1];
	let s2 = st[2] ^ k[2];
	let s3 = st[3] ^ k[3];
	let i = 4;
	for (let r = 1; r < rounds; r++, i += 4) {
		const t0 = mixD(s0, s3, s2, s1, k[i]);
		const t1 = mixD(s1, s0, s3, s2, k[i + 1]);
		const t2 = mixD(s2, s1, s0, s3, k[i + 2]);
		const t3 = mixD(s3, s2, s1, s0, k[i + 3]);
		s0 = t0;
		s1 = t1;
		s2 = t2;
		s3 = t3;
	}
	st[0] = lastRound(INV_SBOX, s0, s3, s2, s1, k[i]);
	st[1] = lastRound(INV_SBOX, s1, s0, s3, s2, k[i + 1]);
	st[2] = lastRound(INV_SBOX, s2, s1, s0, s3, k[i + 2]);
	st[3] = lastRound(INV_SBOX, s3, s2, s1, s0, k[i + 3]);
}

function load(src: Uint8Array, off: number, out: Int32Array): void {
	for (let i = 0; i < 4; i++) {
		const j = off + i * 4;
		out[i] = (src[j] << 24) | (src[j + 1] << 16) | (src[j + 2] << 8) | src[j + 3];
	}
}

function store(words: Int32Array, out: Uint8Array, off: number): void {
	for (let i = 0; i < 4; i++) {
		const w = words[i];
		out[off + i * 4] = w >>> 24;
		out[off + i * 4 + 1] = w >>> 16;
		out[off + i * 4 + 2] = w >>> 8;
		out[off + i * 4 + 3] = w;
	}
}

/** AES with CBC chaining over whole blocks (`data.length` must be a multiple of 16). */
export class SoftwareAes {
	private readonly schedule: ReturnType<typeof expandKey>;

	constructor(key: Uint8Array) {
		initTables();
		this.schedule = expandKey(key);
	}

	encryptCbc(iv: Uint8Array, data: Uint8Array): Uint8Array {
		const { rounds, enc } = this.schedule;
		const out = new Uint8Array(data.length);
		const chain = new Int32Array(4);
		const block = new Int32Array(4);
		load(iv, 0, chain);
		for (let off = 0; off < data.length; off += 16) {
			load(data, off, block);
			for (let i = 0; i < 4; i++) {
				chain[i] ^= block[i];
			}
			encryptBlock(chain, enc, rounds);
			store(chain, out, off);
		}
		return out;
	}

	decryptCbc(iv: Uint8Array, data: Uint8Array): Uint8Array {
		const { rounds, dec } = this.schedule;
		const out = new Uint8Array(data.length);
		const prev = new Int32Array(4);
		const block = new Int32Array(4);
		const cipher = new Int32Array(4);
		load(iv, 0, prev);
		for (let off = 0; off < data.length; off += 16) {
			load(data, off, block);
			cipher.set(block);
			decryptBlock(block, dec, rounds);
			for (let i = 0; i < 4; i++) {
				block[i] ^= prev[i];
			}
			prev.set(cipher);
			store(block, out, off);
		}
		return out;
	}
}
