import { describe, expect, it } from 'vitest';

import type { SupportedLocale } from './locales';
import { createTranslator, interpolate } from './translator';
import type { TranslationKey } from './translations-en';

describe('interpolate', () => {
	it('fills every {name} placeholder from the variables', () => {
		expect(interpolate('{count} of {total} ({count})', { count: 3, total: 9 })).toBe('3 of 9 (3)');
	});

	it('leaves unknown placeholders and text without variables untouched', () => {
		expect(interpolate('Layout {n} of {total}', { n: 2 })).toBe('Layout 2 of {total}');
		expect(interpolate('Style {n}')).toBe('Style {n}');
	});

	it('inserts values literally, without replacement patterns', () => {
		expect(interpolate("'{name}' is protected.", { name: '$&$1 Q1' })).toBe(
			"'$&$1 Q1' is protected.",
		);
	});
});

describe('createTranslator', () => {
	it('resolves English by default', () => {
		const t = createTranslator();
		expect(t('tab.insert')).toBe('Insert');
		expect(t('chart.layoutN', { n: 4 })).toBe('Layout 4');
	});

	it('resolves a built-in locale and interpolates its strings', () => {
		const t = createTranslator({}, 'de');
		expect(t('tab.insert')).toBe('Einfügen');
		expect(t('message.invalidReferenceAt', { ref: 'Tabelle2!A1' })).toBe(
			'Der Bezug ist ungültig: Tabelle2!A1',
		);
		expect(createTranslator({}, 'zh-CN')('tab.insert')).toBe('插入');
	});

	it('prefers overrides, then the locale dictionary, then English', () => {
		const t = createTranslator({ 'tab.insert': 'Hinzufügen' }, 'de');
		expect(t('tab.insert')).toBe('Hinzufügen');
		expect(t('tab.home')).toBe('Start');
		expect(createTranslator({ 'tab.insert': 'Add' })('tab.home')).toBe('Home');
	});

	it('falls back to English for an unknown locale passed at runtime', () => {
		const t = createTranslator({}, 'xx' as SupportedLocale);
		expect(t('tab.insert')).toBe('Insert');
	});

	it('interpolates overrides too', () => {
		const t = createTranslator({ 'pageLayout.pages': '{count} Seiten' });
		expect(t('pageLayout.pages', { count: 3 })).toBe('3 Seiten');
	});

	it('returns the key itself for an unknown key', () => {
		const t = createTranslator();
		expect(t('no.such.key' as TranslationKey)).toBe('no.such.key');
	});
});
