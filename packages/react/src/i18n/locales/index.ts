import { en } from '../translations-en';
import { de } from './de';
import { es } from './es';
import { fr } from './fr';
import type { LocaleInfo, SupportedLocale, TranslationDictionary } from './types';
import { zhCN } from './zh-CN';

export type { AreaDictionary, LocaleInfo, SupportedLocale, TranslationDictionary } from './types';
export { de, es, fr, zhCN };

/** The built-in UI languages, with their native names, in picker order. */
export const BUILT_IN_LOCALES: readonly LocaleInfo[] = [
	{ code: 'en', nativeName: 'English' },
	{ code: 'de', nativeName: 'Deutsch' },
	{ code: 'fr', nativeName: 'Français' },
	{ code: 'es', nativeName: 'Español' },
	{ code: 'zh-CN', nativeName: '简体中文' },
];

/** The complete dictionary of every built-in locale. */
export const BUILT_IN_DICTIONARIES: Readonly<Record<SupportedLocale, TranslationDictionary>> = {
	en,
	de,
	fr,
	es,
	'zh-CN': zhCN,
};

/** `value` names a built-in locale (`'de'`, `'zh-CN'`, ...). */
export function isSupportedLocale(value: unknown): value is SupportedLocale {
	return typeof value === 'string' && BUILT_IN_LOCALES.some((l) => l.code === value);
}
