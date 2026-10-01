import type { Workbook } from 'xlsx-viewer-core';

const cache = new WeakMap<Workbook, Map<string, string>>();

/** Object URL for an image part of the workbook (created once per workbook and path). */
export function mediaUrl(workbook: Workbook, path: string): string | undefined {
	let urls = cache.get(workbook);
	if (!urls) {
		urls = new Map();
		cache.set(workbook, urls);
	}
	const hit = urls.get(path);
	if (hit) {
		return hit;
	}
	const media = workbook.media.get(path);
	if (!media || typeof URL === 'undefined' || typeof URL.createObjectURL !== 'function') {
		return undefined;
	}
	const url = URL.createObjectURL(new Blob([media.data as BlobPart], { type: media.contentType }));
	urls.set(path, url);
	return url;
}

/** Whether the browser can display an image content type natively. */
export function isDisplayableImage(contentType: string): boolean {
	return /^image\/(png|jpe?g|gif|bmp|webp|svg\+xml|avif)$/i.test(contentType);
}
