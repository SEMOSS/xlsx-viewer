import { imageContentType } from '../package/content-types';
import { partExtension } from '../package/paths';
import { generateId } from '../sheet/worksheet';
import type { ReadContext } from './context';
import type { BlipResolver } from './drawingml';
import type { RelLookup } from './part-io';

/** Load a binary part (image) into `ctx.media`. Resolves to `false` when the part is missing. */
export async function loadMedia(ctx: ReadContext, path: string): Promise<boolean> {
	if (ctx.media.has(path)) {
		return true;
	}
	const data = await ctx.pkg.bytes(path);
	if (!data) {
		return false;
	}
	const contentType = ctx.pkg.contentTypes.get(path) ?? imageContentType(partExtension(path));
	ctx.media.set(path, { path, contentType, data });
	return true;
}

/** Per-part state while one drawing part is read: relationships, ids in use, pending image loads. */
export class DrawingReader {
	private readonly ids = new Set<string>();
	private readonly pending: Array<Promise<boolean>> = [];
	/** Synchronous `r:embed` resolver for picture fills; the image loads in the background. */
	readonly blipResolver: BlipResolver;

	constructor(
		readonly ctx: ReadContext,
		readonly partPath: string,
		readonly rels: RelLookup,
		readonly sheetName: string,
	) {
		this.blipResolver = (relId) => {
			const path = this.rels.path(relId);
			if (path) {
				this.pending.push(loadMedia(ctx, path));
			}
			return path;
		};
	}

	/** The `cNvPr` id when it is unused in this drawing, else a generated one. */
	uniqueId(id: string | undefined): string {
		let out = id && id.trim() ? id.trim() : '';
		if (!out || this.ids.has(out)) {
			out = generateId('drawing');
		}
		this.ids.add(out);
		return out;
	}

	/** Resolve and load the image of an `r:embed` / `r:link`; resolves to its media path. */
	async loadImage(relId: string | undefined): Promise<string | undefined> {
		const path = this.rels.path(relId);
		if (!path) {
			return undefined;
		}
		return (await loadMedia(this.ctx, path)) ? path : undefined;
	}

	warn(code: string, message: string, part = this.partPath): void {
		this.ctx.warn(code, message, { part, sheet: this.sheetName });
	}

	/** Wait for background image loads started by {@link blipResolver}. */
	async settle(): Promise<void> {
		await Promise.all(this.pending);
	}
}
