import type { HeaderFooter } from '../types/sheet';
import type { ColorSpec } from '../types/style';
import type { HeaderFooterFont, HeaderFooterRun, HeaderFooterSections } from './types';

/**
 * Excel header / footer codes: `&L` `&C` `&R` sections, fields (`&P`,
 * `&P+n`, `&N`, `&D`, `&T`, `&F`, `&A`, `&Z`), `&&`, fonts (`&"Name,Style"`,
 * `&nn`), toggles (`&B &I &U &E &S &X &Y`) and colors (`&KRRGGBB`,
 * `&KTTSNNN` for a theme color with a tint). Pictures (`&G`) are skipped.
 */

export interface HeaderFooterContext {
	/** Page number (`&P`). */
	page: number;
	/** Total pages (`&N`). */
	pages: number;
	/** Sheet tab name (`&A`). */
	sheetName: string;
	/** File name (`&F`). */
	fileName?: string;
	/** Full path (`&Z`). */
	filePath?: string;
	/** Date and time for `&D` / `&T` (default: now). */
	date?: Date;
}

type SectionName = keyof HeaderFooterSections;

/** `M/D/YYYY`, Excel's default short date (en-US). */
export function formatHeaderDate(d: Date): string {
	return `${d.getMonth() + 1}/${d.getDate()}/${d.getFullYear()}`;
}

/** `h:mm AM/PM`, Excel's default time (en-US). */
export function formatHeaderTime(d: Date): string {
	const h = d.getHours();
	const m = String(d.getMinutes()).padStart(2, '0');
	return `${h % 12 === 0 ? 12 : h % 12}:${m} ${h < 12 ? 'AM' : 'PM'}`;
}

/** `&"Arial,Bold Italic"`: `-` keeps the current font name. */
function applyFontSpec(spec: string, font: HeaderFooterFont): HeaderFooterFont {
	const comma = spec.indexOf(',');
	const name = (comma >= 0 ? spec.slice(0, comma) : spec).trim();
	const style = comma >= 0 ? spec.slice(comma + 1).toLowerCase() : '';
	const next: HeaderFooterFont = { ...font };
	if (name && name !== '-') {
		next.name = name;
	}
	if (comma >= 0) {
		next.bold = /bold/.test(style) || undefined;
		next.italic = /italic|oblique/.test(style) || undefined;
	}
	return next;
}

/** `RRGGBB` or `TTSNNN` (theme index, sign, tint percent) after `&K`. */
function parseColor(text: string): ColorSpec | undefined {
	if (/^[0-9A-Fa-f]{6}$/.test(text)) {
		return { rgb: `FF${text.toUpperCase()}` };
	}
	const m = /^(\d\d)([+-])(\d{3})$/.exec(text);
	if (m) {
		const tint = (m[2] === '-' ? -1 : 1) * (Number(m[3]) / 100);
		return tint === 0 ? { theme: Number(m[1]) } : { theme: Number(m[1]), tint };
	}
	return undefined;
}

function toggle(font: HeaderFooterFont, key: 'bold' | 'italic' | 'strike'): HeaderFooterFont {
	return { ...font, [key]: font[key] ? undefined : true };
}

function clean(font: HeaderFooterFont): HeaderFooterFont {
	const out: HeaderFooterFont = {};
	for (const [k, v] of Object.entries(font) as Array<[keyof HeaderFooterFont, unknown]>) {
		if (v !== undefined) {
			(out as Record<string, unknown>)[k] = v;
		}
	}
	return out;
}

/** Format a header / footer string for a page. */
export function formatHeaderFooter(
	code: string | undefined,
	ctx: HeaderFooterContext,
): HeaderFooterSections {
	const out: HeaderFooterSections = { left: [], center: [], right: [] };
	if (!code) {
		return out;
	}
	const date = ctx.date ?? new Date();
	let section: SectionName = 'center';
	let font: HeaderFooterFont = {};
	let text = '';
	const flush = () => {
		if (text) {
			const runs: HeaderFooterRun[] = out[section];
			const last = runs[runs.length - 1];
			const f = clean(font);
			if (last && JSON.stringify(last.font) === JSON.stringify(f)) {
				last.text += text;
			} else {
				runs.push({ text, font: f });
			}
			text = '';
		}
	};
	const setFont = (next: HeaderFooterFont) => {
		flush();
		font = next;
	};
	for (let i = 0; i < code.length; i++) {
		const ch = code[i];
		if (ch !== '&') {
			text += ch;
			continue;
		}
		const c = code[i + 1];
		if (c === undefined) {
			break;
		}
		i++;
		if (/[0-9]/.test(c)) {
			const digits = /^\d{1,3}/.exec(code.slice(i))![0];
			i += digits.length - 1;
			setFont({ ...font, size: Number(digits) });
			continue;
		}
		switch (c.toUpperCase()) {
			case '&':
				text += '&';
				break;
			case 'L':
			case 'C':
			case 'R':
				flush();
				section = c.toUpperCase() === 'L' ? 'left' : c.toUpperCase() === 'C' ? 'center' : 'right';
				font = {};
				break;
			case 'P': {
				const m = /^([+-])(\d+)/.exec(code.slice(i + 1));
				let page = ctx.page;
				if (m) {
					page += (m[1] === '-' ? -1 : 1) * Number(m[2]);
					i += m[0].length;
				}
				text += String(page);
				break;
			}
			case 'N':
				text += String(ctx.pages);
				break;
			case 'D':
				text += formatHeaderDate(date);
				break;
			case 'T':
				text += formatHeaderTime(date);
				break;
			case 'F':
				text += ctx.fileName ?? '';
				break;
			case 'Z':
				text += ctx.filePath ?? '';
				break;
			case 'A':
				text += ctx.sheetName;
				break;
			case 'B':
				setFont(toggle(font, 'bold'));
				break;
			case 'I':
				setFont(toggle(font, 'italic'));
				break;
			case 'S':
				setFont(toggle(font, 'strike'));
				break;
			case 'U':
				setFont({ ...font, underline: font.underline === 'single' ? undefined : 'single' });
				break;
			case 'E':
				setFont({ ...font, underline: font.underline === 'double' ? undefined : 'double' });
				break;
			case 'X':
				setFont({
					...font,
					vertAlign: font.vertAlign === 'superscript' ? undefined : 'superscript',
				});
				break;
			case 'Y':
				setFont({ ...font, vertAlign: font.vertAlign === 'subscript' ? undefined : 'subscript' });
				break;
			case '"': {
				const end = code.indexOf('"', i + 1);
				const stop = end < 0 ? code.length : end;
				setFont(applyFontSpec(code.slice(i + 1, stop), font));
				i = stop;
				break;
			}
			case 'K': {
				const color = parseColor(code.slice(i + 1, i + 7));
				if (color) {
					setFont({ ...font, color });
					i += 6;
				}
				break;
			}
			default:
				// &G (picture), &O (outline), &H (shadow) and unknown codes print nothing.
				break;
		}
	}
	flush();
	return out;
}

/** Plain text of formatted sections (runs joined), for previews and tests. */
export function headerFooterText(s: HeaderFooterSections): {
	left: string;
	center: string;
	right: string;
} {
	const join = (runs: HeaderFooterRun[]) => runs.map((r) => r.text).join('');
	return { left: join(s.left), center: join(s.center), right: join(s.right) };
}

/** The raw codes of a header / footer split into its three sections (for editing). */
export function splitHeaderFooter(code: string | undefined): {
	left: string;
	center: string;
	right: string;
} {
	const parts = { left: '', center: '', right: '' };
	if (!code) {
		return parts;
	}
	let section: SectionName = 'center';
	let start = 0;
	for (let i = 0; i < code.length; i++) {
		if (code[i] !== '&') {
			continue;
		}
		const c = code[i + 1]?.toUpperCase();
		if (c === '"') {
			const end = code.indexOf('"', i + 2);
			i = end < 0 ? code.length : end;
		} else if (c === 'L' || c === 'C' || c === 'R') {
			parts[section] += code.slice(start, i);
			section = c === 'L' ? 'left' : c === 'C' ? 'center' : 'right';
			start = i + 2;
			i++;
		} else {
			i++;
		}
	}
	parts[section] += code.slice(start);
	return parts;
}

/** Join edited sections back into a header / footer string (`''` when all are empty). */
export function joinHeaderFooter(parts: {
	left?: string;
	center?: string;
	right?: string;
}): string {
	return (
		(parts.left ? `&L${parts.left}` : '') +
		(parts.center ? `&C${parts.center}` : '') +
		(parts.right ? `&R${parts.right}` : '')
	);
}

/**
 * Header and footer codes for a page: the first page uses the first-page
 * variant with `differentFirst`, even page numbers the even variant with
 * `differentOddEven`.
 */
export function headerFooterFor(
	hf: HeaderFooter | undefined,
	pageIndex: number,
	pageNumber: number,
): { header?: string; footer?: string } {
	if (!hf) {
		return {};
	}
	if (hf.differentFirst && pageIndex === 0) {
		return { header: hf.firstHeader, footer: hf.firstFooter };
	}
	if (hf.differentOddEven && pageNumber % 2 === 0) {
		return { header: hf.evenHeader, footer: hf.evenFooter };
	}
	return { header: hf.oddHeader, footer: hf.oddFooter };
}

/** The sections have no text at all. */
export function isHeaderFooterEmpty(s: HeaderFooterSections | null | undefined): boolean {
	return !s || (s.left.length === 0 && s.center.length === 0 && s.right.length === 0);
}
