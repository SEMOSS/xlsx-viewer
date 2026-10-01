import { isSupportedLocale, type SupportedLocale } from 'xlsx-react-viewer';

const STORAGE_KEY = 'xlsx-demo-locale';

/** The first built-in locale matching the browser languages (`de-AT` -> `de`, `zh-Hans` -> `zh-CN`). */
function matchBrowserLocale(languages: readonly string[]): SupportedLocale {
	for (const tag of languages) {
		const lower = tag.toLowerCase();
		if (lower.startsWith('zh')) {
			// Simplified Chinese only: Traditional variants (zh-TW, zh-HK, zh-Hant) keep looking.
			if (/^zh(-(cn|sg|hans)(-|$)|$)/.test(lower)) {
				return 'zh-CN';
			}
			continue;
		}
		const base = lower.split('-')[0];
		if (isSupportedLocale(base)) {
			return base;
		}
	}
	return 'en';
}

/** The saved UI language, else the browser's, else English. */
export function readLocale(): SupportedLocale {
	try {
		const stored = localStorage.getItem(STORAGE_KEY);
		if (isSupportedLocale(stored)) {
			return stored;
		}
	} catch {
		/* ignore */
	}
	const languages = navigator.languages?.length ? navigator.languages : [navigator.language];
	return matchBrowserLocale(languages.filter(Boolean));
}

/** Remember the UI language (like the theme). */
export function saveLocale(locale: SupportedLocale): void {
	try {
		localStorage.setItem(STORAGE_KEY, locale);
	} catch {
		/* ignore */
	}
}
