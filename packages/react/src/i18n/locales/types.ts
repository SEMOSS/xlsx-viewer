import type { TranslationKey } from '../translations-en';

/** The UI languages that ship with the viewer. */
export type SupportedLocale = 'en' | 'de' | 'fr' | 'es' | 'zh-CN';

/** A built-in locale and its name in its own language (for language pickers). */
export interface LocaleInfo {
	code: SupportedLocale;
	nativeName: string;
}

/** A complete dictionary: every UI string, keyed like the English one. */
export type TranslationDictionary = Readonly<Record<TranslationKey, string>>;

/**
 * One area file of a dictionary, keyed exactly like its English counterpart:
 * a missing or misspelled key is a type error.
 */
export type AreaDictionary<T> = { readonly [K in keyof T]: string };
