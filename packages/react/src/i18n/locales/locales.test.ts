import { describe, expect, it } from 'vitest';

import { en, EN_AREAS, type TranslationKey } from '../translations-en';
import { BUILT_IN_DICTIONARIES, BUILT_IN_LOCALES, isSupportedLocale } from '.';

const KEYS = Object.keys(en).sort() as TranslationKey[];
const EM_DASH = '\u2014';

/** The `{name}` placeholders of a string, sorted. */
function placeholders(text: string): string[] {
	return [...text.matchAll(/\{(\w+)\}/g)].map((m) => m[1]).sort();
}

/** Keyboard shortcut texts (`Ctrl+S`, `Alt+Q`, `Shift+F2`, `F9`), which stay as written. */
function shortcuts(text: string): string[] {
	return [...text.matchAll(/\b(?:(?:Ctrl|Alt|Shift)\+)+\S+?(?=\)|$|\s)|\bF\d{1,2}\b/g)]
		.map((m) => m[0])
		.sort();
}

describe('built-in locales', () => {
	it('lists each supported locale once, English first', () => {
		const codes = BUILT_IN_LOCALES.map((l) => l.code);
		expect(codes).toEqual(['en', 'de', 'fr', 'es', 'zh-CN']);
		expect(Object.keys(BUILT_IN_DICTIONARIES).sort()).toEqual([...codes].sort());
		expect(BUILT_IN_LOCALES.every((l) => l.nativeName.trim().length > 0)).toBe(true);
		expect(isSupportedLocale('zh-CN')).toBe(true);
		expect(isSupportedLocale('zh')).toBe(false);
	});

	it('defines every English key in exactly one area file', () => {
		const counts = new Map<string, number>();
		for (const area of Object.values(EN_AREAS)) {
			for (const key of Object.keys(area)) {
				counts.set(key, (counts.get(key) ?? 0) + 1);
			}
		}
		expect([...counts].filter(([, n]) => n > 1).map(([key]) => key)).toEqual([]);
		expect(counts.size).toBe(KEYS.length);
	});

	it('uses keys made of an area prefix and camelCase segments', () => {
		expect(KEYS.filter((key) => !/^[a-z][a-zA-Z0-9]*(\.[a-zA-Z0-9]+)+$/.test(key))).toEqual([]);
	});
});

describe.each(BUILT_IN_LOCALES.map((l) => l.code))('the %s dictionary', (code) => {
	const dictionary = BUILT_IN_DICTIONARIES[code];

	it('has exactly the English key set', () => {
		expect(Object.keys(dictionary).sort()).toEqual(KEYS);
	});

	it('has no empty value', () => {
		expect(KEYS.filter((key) => !dictionary[key]?.trim())).toEqual([]);
	});

	it('keeps the {placeholders} of the English text', () => {
		const broken = KEYS.filter(
			(key) => placeholders(dictionary[key] ?? '').join() !== placeholders(en[key]).join(),
		);
		expect(broken).toEqual([]);
	});

	it('keeps keyboard shortcut texts as written', () => {
		const changed = KEYS.filter(
			(key) => shortcuts(dictionary[key] ?? '').join() !== shortcuts(en[key]).join(),
		);
		expect(changed).toEqual([]);
	});

	it('marks the same commands with a trailing ellipsis', () => {
		const mismatched = KEYS.filter(
			(key) => en[key].endsWith('...') !== (dictionary[key] ?? '').endsWith('...'),
		);
		expect(mismatched).toEqual([]);
	});

	it('contains no em-dash', () => {
		expect(KEYS.filter((key) => (dictionary[key] ?? '').includes(EM_DASH))).toEqual([]);
	});
});

describe('Microsoft Office terminology', () => {
	const TERMS: Array<[TranslationKey, Record<'de' | 'fr' | 'es' | 'zh-CN', string>]> = [
		['tab.file', { de: 'Datei', fr: 'Fichier', es: 'Archivo', 'zh-CN': '文件' }],
		['tab.home', { de: 'Start', fr: 'Accueil', es: 'Inicio', 'zh-CN': '开始' }],
		['tab.insert', { de: 'Einfügen', fr: 'Insertion', es: 'Insertar', 'zh-CN': '插入' }],
		[
			'tab.pageLayout',
			{ de: 'Seitenlayout', fr: 'Mise en page', es: 'Disposición de página', 'zh-CN': '页面布局' },
		],
		['tab.formulas', { de: 'Formeln', fr: 'Formules', es: 'Fórmulas', 'zh-CN': '公式' }],
		['tab.data', { de: 'Daten', fr: 'Données', es: 'Datos', 'zh-CN': '数据' }],
		['tab.review', { de: 'Überprüfen', fr: 'Révision', es: 'Revisar', 'zh-CN': '审阅' }],
		['tab.view', { de: 'Ansicht', fr: 'Affichage', es: 'Vista', 'zh-CN': '视图' }],
		[
			'dialog.formatCells.title',
			{
				de: 'Zellen formatieren',
				fr: 'Format de cellule',
				es: 'Formato de celdas',
				'zh-CN': '设置单元格格式',
			},
		],
		[
			'cmd.conditionalFormatting',
			{
				de: 'Bedingte Formatierung',
				fr: 'Mise en forme conditionnelle',
				es: 'Formato condicional',
				'zh-CN': '条件格式',
			},
		],
	];

	it.each(TERMS)('translates %s as localized Office does', (key, expected) => {
		for (const [code, text] of Object.entries(expected)) {
			expect(BUILT_IN_DICTIONARIES[code as keyof typeof expected][key]).toBe(text);
		}
	});
});
