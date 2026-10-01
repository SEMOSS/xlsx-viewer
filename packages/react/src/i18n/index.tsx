import { createContext, useContext, useMemo, type ReactNode } from 'react';

import type { SupportedLocale } from './locales';
import { createTranslator, type Translate } from './translator';
import type { Translations } from './translations-en';

export {
	BUILT_IN_DICTIONARIES,
	BUILT_IN_LOCALES,
	de,
	es,
	fr,
	isSupportedLocale,
	zhCN,
	type AreaDictionary,
	type LocaleInfo,
	type SupportedLocale,
	type TranslationDictionary,
} from './locales';
export { FUNCTION_CATEGORY_KEYS } from './function-categories';
export { formatNumber, formatPercent, numberFormatter } from './number-format';
export { createTranslator, interpolate, type Translate, type TranslationVars } from './translator';
export { en, type TranslationKey, type Translations } from './translations-en';

interface I18nValue {
	t: Translate;
	locale: SupportedLocale;
}

const I18nContext = createContext<I18nValue>({ t: createTranslator(), locale: 'en' });

export interface I18nProviderProps {
	/** Built-in dictionary to use (default `'en'`). */
	locale?: SupportedLocale;
	/** Per-key overrides applied on top of the locale's dictionary. */
	translations?: Translations;
	children: ReactNode;
}

/** Provides the translate function and locale to the viewer below it. */
export function I18nProvider({ locale = 'en', translations, children }: I18nProviderProps) {
	const value = useMemo(
		() => ({ t: createTranslator(translations, locale), locale }),
		[translations, locale],
	);
	return <I18nContext.Provider value={value}>{children}</I18nContext.Provider>;
}

/** The translate function of the enclosing viewer (English outside a provider). */
export function useT(): Translate {
	return useContext(I18nContext).t;
}

/** The UI locale of the enclosing viewer, for `Intl` number formatting in the chrome. */
export function useLocale(): SupportedLocale {
	return useContext(I18nContext).locale;
}
