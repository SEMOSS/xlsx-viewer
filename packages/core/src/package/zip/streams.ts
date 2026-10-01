/**
 * Raw DEFLATE through the platform's CompressionStream / DecompressionStream
 * (browsers, Node 21+, Deno, Bun). Much faster than a JavaScript inflater;
 * callers fall back to JSZip where the streams are missing.
 */

let native: boolean | undefined;

/** The platform has `deflate-raw` compression streams. */
export function hasNativeDeflate(): boolean {
	if (native === undefined) {
		try {
			native =
				typeof CompressionStream !== 'undefined' &&
				typeof DecompressionStream !== 'undefined' &&
				new CompressionStream('deflate-raw') !== null &&
				new DecompressionStream('deflate-raw') !== null;
		} catch {
			native = false;
		}
	}
	return native;
}

/** Test hook: force the JSZip fallback (`false`) or re-detect (`undefined`). */
export function setNativeDeflateForTests(value: boolean | undefined): void {
	native = value;
}

async function pipe(
	data: Uint8Array,
	transform: { writable: WritableStream<BufferSource>; readable: ReadableStream<Uint8Array> },
): Promise<Uint8Array> {
	const writer = transform.writable.getWriter();
	const written = writer.write(data as BufferSource).then(() => writer.close());
	const reader = transform.readable.getReader();
	const chunks: Uint8Array[] = [];
	let total = 0;
	for (;;) {
		const { done, value } = await reader.read();
		if (done) {
			break;
		}
		chunks.push(value);
		total += value.length;
	}
	await written;
	if (chunks.length === 1) {
		return chunks[0];
	}
	const out = new Uint8Array(total);
	let offset = 0;
	for (const c of chunks) {
		out.set(c, offset);
		offset += c.length;
	}
	return out;
}

export function inflateRaw(data: Uint8Array): Promise<Uint8Array> {
	return pipe(data, new DecompressionStream('deflate-raw'));
}

export function deflateRaw(data: Uint8Array): Promise<Uint8Array> {
	return pipe(data, new CompressionStream('deflate-raw'));
}
