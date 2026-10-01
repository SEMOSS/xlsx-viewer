import { createDefaultTheme } from '../styles/defaults';
import type { Theme, ThemeColorScheme } from '../types/theme';
import { parseXml } from '../xml/parser';
import type { XmlElement } from '../xml/types';
import { firstElement, kid, lname, str } from './xml-utils';

/** Well-known `sysClr` values when `lastClr` is missing. */
const SYSTEM_COLORS: Record<string, string> = {
	windowText: '000000',
	window: 'FFFFFF',
	buttonFace: 'F0F0F0',
	buttonText: '000000',
	highlight: '0078D7',
	highlightText: 'FFFFFF',
	grayText: '6D6D6D',
	captionText: '000000',
	menuText: '000000',
};

/** Basic `prstClr` names that can appear in a color scheme. */
const PRESET_COLORS: Record<string, string> = {
	black: '000000',
	white: 'FFFFFF',
	red: 'FF0000',
	green: '008000',
	blue: '0000FF',
	yellow: 'FFFF00',
	gray: '808080',
};

const SLOTS = [
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
] as const;

function slotColor(slot: XmlElement | undefined): string | undefined {
	const c = firstElement(slot);
	if (!c) {
		return undefined;
	}
	let hex: string | undefined;
	switch (lname(c)) {
		case 'srgbClr':
			hex = str(c, 'val');
			break;
		case 'sysClr':
			hex = str(c, 'lastClr') ?? SYSTEM_COLORS[str(c, 'val') ?? ''];
			break;
		case 'prstClr':
			hex = PRESET_COLORS[str(c, 'val') ?? ''];
			break;
		default:
			hex = undefined;
	}
	return hex && /^[0-9A-Fa-f]{6}$/.test(hex) ? hex.toUpperCase() : undefined;
}

function typeface(el: XmlElement | undefined): string | undefined {
	const v = str(el, 'typeface');
	return v ? v : undefined;
}

/**
 * `xl/theme/theme1.xml` -> {@link Theme}: the color scheme, the major / minor
 * fonts and the original XML (re-emitted on save). Missing pieces fall back
 * to the default Office theme.
 */
export function parseTheme(xml: string): Theme {
	const fallback = createDefaultTheme();
	const root = parseXml(xml).root;
	const elements = kid(root, 'themeElements');
	const scheme = kid(elements, 'clrScheme');
	const colors: ThemeColorScheme = {
		...fallback.colors,
		name: str(scheme, 'name') ?? fallback.colors.name,
	};
	for (const slot of SLOTS) {
		const hex = slotColor(kid(scheme, slot));
		if (hex) {
			colors[slot] = hex;
		}
	}
	const fontScheme = kid(elements, 'fontScheme');
	const major = kid(fontScheme, 'majorFont');
	const minor = kid(fontScheme, 'minorFont');
	const fonts = {
		name: str(fontScheme, 'name') ?? fallback.fonts.name,
		majorLatin: typeface(kid(major, 'latin')) ?? fallback.fonts.majorLatin,
		minorLatin: typeface(kid(minor, 'latin')) ?? fallback.fonts.minorLatin,
	} as Theme['fonts'];
	const extras: Array<[keyof Theme['fonts'], string | undefined]> = [
		['majorEastAsian', typeface(kid(major, 'ea'))],
		['minorEastAsian', typeface(kid(minor, 'ea'))],
		['majorComplexScript', typeface(kid(major, 'cs'))],
		['minorComplexScript', typeface(kid(minor, 'cs'))],
	];
	for (const [key, value] of extras) {
		if (value) {
			fonts[key] = value;
		}
	}
	return { name: str(root, 'name') ?? fallback.name, colors, fonts, rawXml: xml };
}
