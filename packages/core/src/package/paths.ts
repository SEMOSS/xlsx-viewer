/**
 * OPC part-name helpers. Part paths inside the package are stored WITHOUT a
 * leading slash (`xl/worksheets/sheet1.xml`), the way JSZip names entries.
 */

/** Normalise `a/b/../c` and `./x`, strip any leading slash. */
export function normalizePartPath(p: string): string {
	const segs: string[] = [];
	for (const s of p.replace(/\\/g, '/').split('/')) {
		if (s === '' || s === '.') {
			continue;
		}
		if (s === '..') {
			segs.pop();
		} else {
			segs.push(s);
		}
	}
	return segs.join('/');
}

/** Directory of a part path (`xl/worksheets/sheet1.xml` -> `xl/worksheets`). */
export function partDirectory(p: string): string {
	const i = p.lastIndexOf('/');
	return i === -1 ? '' : p.slice(0, i);
}

/** File name of a part path. */
export function partFileName(p: string): string {
	const i = p.lastIndexOf('/');
	return i === -1 ? p : p.slice(i + 1);
}

/** Path of the relationships part for a source part (`xl/workbook.xml` -> `xl/_rels/workbook.xml.rels`). */
export function relsPathFor(partPath: string): string {
	const dir = partDirectory(partPath);
	const file = partFileName(partPath);
	return dir ? `${dir}/_rels/${file}.rels` : `_rels/${file}.rels`;
}

/**
 * Resolve a relationship target relative to its source part. Absolute
 * targets (`/xl/media/image1.png`) are taken from the package root.
 */
export function resolveRelTarget(sourcePart: string, target: string): string {
	if (target.startsWith('/')) {
		return normalizePartPath(target);
	}
	const base = partDirectory(sourcePart);
	return normalizePartPath(base ? `${base}/${target}` : target);
}

/** Relative target from `sourcePart` to `targetPart` (both package paths). */
export function relativeTarget(sourcePart: string, targetPart: string): string {
	const from = partDirectory(sourcePart).split('/').filter(Boolean);
	const to = normalizePartPath(targetPart).split('/');
	let i = 0;
	while (i < from.length && i < to.length - 1 && from[i] === to[i]) {
		i++;
	}
	const ups = from.length - i;
	return [...Array(ups).fill('..'), ...to.slice(i)].join('/');
}

/** Lower-case extension without the dot. */
export function partExtension(p: string): string {
	const file = partFileName(p);
	const i = file.lastIndexOf('.');
	return i === -1 ? '' : file.slice(i + 1).toLowerCase();
}
