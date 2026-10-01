import { interpolate, type Translate } from '../../i18n/translator';
import type { TranslationKey } from '../../i18n/translations-en';

/**
 * Excel's built-in headers / footers (Page Setup > Header/Footer). Like
 * Excel's localized presets, the words they write into the sheet (Page, of,
 * Confidential) follow the UI language; the field codes (`&P`, `&N`, `&D`,
 * `&F`, `&A`) do not.
 */

/** This workbook's names and page 1, as the preset list shows them. */
export interface PresetLabels {
	sheet: string;
	file: string;
	date: string;
	/** The number 1 in the UI locale ("Page 1"). */
	page: string;
}

export interface HeaderFooterPreset {
	/** Header / footer code written into the sheet ('' for (none)). */
	code: string;
	/** Caption in the preset list. */
	label: string;
}

type Section = 'left' | 'center' | 'right';

/** What a section of a built-in header / footer shows. */
type PresetField = 'page' | 'pageOf' | 'sheet' | 'file' | 'date' | 'confidential';

/** The presets in Excel's order, (none) first: the field of each section used. */
const PRESETS: ReadonlyArray<Partial<Record<Section, PresetField>>> = [
	{},
	{ center: 'page' },
	{ center: 'pageOf' },
	{ center: 'sheet' },
	{ left: 'confidential', center: 'date', right: 'page' },
	{ center: 'file' },
	{ center: 'sheet', right: 'page' },
	{ center: 'file', right: 'page' },
	{ center: 'page', right: 'sheet' },
	{ center: 'page', right: 'file' },
];

const SECTION_CODES: ReadonlyArray<[Section, string]> = [
	['left', '&L'],
	['center', '&C'],
	['right', '&R'],
];

/** Translated words as header / footer code: `&` doubled, then the field codes filled in. */
function codeText(t: Translate, key: TranslationKey, codes?: Record<string, string>): string {
	return interpolate(t(key).replace(/&/g, '&&'), codes);
}

function fieldCode(field: PresetField, t: Translate): string {
	switch (field) {
		case 'page':
			return codeText(t, 'dialog.pageSetup.presetPage', { page: '&P' });
		case 'pageOf':
			return codeText(t, 'dialog.pageSetup.presetPageOf', { page: '&P', pages: '&N' });
		case 'sheet':
			return '&A';
		case 'file':
			return '&F';
		case 'date':
			return '&D';
		case 'confidential':
			return codeText(t, 'dialog.pageSetup.presetConfidential');
	}
}

function fieldLabel(field: PresetField, t: Translate, labels: PresetLabels): string {
	switch (field) {
		case 'page':
			return t('dialog.pageSetup.presetPage', { page: labels.page });
		case 'pageOf':
			return t('dialog.pageSetup.presetPageOf', { page: labels.page, pages: '?' });
		case 'sheet':
			return labels.sheet;
		case 'file':
			return labels.file;
		case 'date':
			return labels.date;
		case 'confidential':
			return t('dialog.pageSetup.presetConfidential');
	}
}

/** The built-in headers / footers in the language of `t`, captions shown with `labels`. */
export function headerFooterPresets(t: Translate, labels: PresetLabels): HeaderFooterPreset[] {
	return PRESETS.map((preset) => {
		const used = SECTION_CODES.flatMap(([section, code]) => {
			const field = preset[section];
			return field ? [{ code, field }] : [];
		});
		if (used.length === 0) {
			return { code: '', label: t('dialog.pageSetup.presetNone') };
		}
		return {
			code: used.map((u) => u.code + fieldCode(u.field, t)).join(''),
			label: used.map((u) => fieldLabel(u.field, t, labels)).join(', '),
		};
	});
}
