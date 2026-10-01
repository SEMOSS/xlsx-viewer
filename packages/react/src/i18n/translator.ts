import { BUILT_IN_DICTIONARIES, type SupportedLocale } from './locales';
import { en, type TranslationKey, type Translations } from './translations-en';

/** Variables substituted into `{name}` placeholders. */
export type TranslationVars = Readonly<Record<string, string | number>>;

/** Look up a UI string by key, filling `{name}` placeholders from `vars`. */
export type Translate = (key: TranslationKey, vars?: TranslationVars) => string;

const PLACEHOLDER = /\{(\w+)\}/g;

/**
 * Replace `{name}` placeholders with `vars[name]`. Unknown placeholders stay
 * as written, and values are inserted literally (no `$&` patterns).
 */
export function interpolate(text: string, vars?: TranslationVars): string {
	if (!vars) {
		return text;
	}
	return text.replace(PLACEHOLDER, (match, name: string) =>
		Object.prototype.hasOwnProperty.call(vars, name) ? String(vars[name]) : match,
	);
}

/**
 * Build a translate function. Keys resolve from `overrides` first, then the
 * built-in dictionary of `locale`, then English (and finally the key itself).
 */
export function createTranslator(
	overrides: Translations = {},
	locale: SupportedLocale = 'en',
): Translate {
	const dictionary: Translations = BUILT_IN_DICTIONARIES[locale] ?? en;
	return (key, vars) => interpolate(overrides[key] ?? dictionary[key] ?? en[key] ?? key, vars);
}
