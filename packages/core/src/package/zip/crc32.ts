/** CRC-32 (IEEE 802.3), as ZIP stores it for every entry. */

const TABLE = (() => {
	const t = new Uint32Array(256 * 4);
	for (let n = 0; n < 256; n++) {
		let c = n;
		for (let k = 0; k < 8; k++) {
			c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
		}
		t[n] = c >>> 0;
	}
	// Slicing-by-4 tables.
	for (let n = 0; n < 256; n++) {
		let c = t[n];
		for (let k = 1; k < 4; k++) {
			c = t[c & 0xff] ^ (c >>> 8);
			t[k * 256 + n] = c >>> 0;
		}
	}
	return t;
})();

export function crc32(data: Uint8Array): number {
	let c = 0xffffffff;
	let i = 0;
	const n = data.length;
	for (; i + 4 <= n; i += 4) {
		c ^= data[i] | (data[i + 1] << 8) | (data[i + 2] << 16) | (data[i + 3] << 24);
		c =
			TABLE[768 + (c & 0xff)] ^
			TABLE[512 + ((c >>> 8) & 0xff)] ^
			TABLE[256 + ((c >>> 16) & 0xff)] ^
			TABLE[(c >>> 24) & 0xff];
	}
	for (; i < n; i++) {
		c = TABLE[(c ^ data[i]) & 0xff] ^ (c >>> 8);
	}
	return (c ^ 0xffffffff) >>> 0;
}
