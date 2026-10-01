import type { Hyperlink } from '../types/sheet';

/** Schemes a browser may open from a workbook link (never `javascript:`, `data:` and the like). */
const SAFE_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'ftp:', 'tel:', 'news:']);

/**
 * Where a hyperlink points: an external URL (with `#location` appended when
 * both are set) or an internal location (`Sheet2!A1`, a defined name).
 */
export function resolveHyperlink(
	link: Hyperlink,
):
	| { kind: 'url'; url: string }
	| { kind: 'location'; location: string }
	| { kind: 'unsupported'; target: string }
	| null {
	const target = link.target?.trim();
	if (target) {
		const url = link.location ? `${target}#${link.location}` : target;
		let scheme: string;
		try {
			scheme = new URL(url).protocol.toLowerCase();
		} catch {
			// Relative paths (other files next to the workbook) cannot be opened from a browser.
			return { kind: 'unsupported', target };
		}
		return SAFE_SCHEMES.has(scheme) ? { kind: 'url', url } : { kind: 'unsupported', target };
	}
	const location = link.location?.trim().replace(/^#/, '');
	return location ? { kind: 'location', location } : null;
}

/** Text for a link tooltip: its screen tip, else where it goes. */
export function hyperlinkDescription(link: Hyperlink): string {
	if (link.tooltip) {
		return link.tooltip;
	}
	if (link.target) {
		return link.location ? `${link.target} - ${link.location}` : link.target;
	}
	return link.location ?? '';
}
