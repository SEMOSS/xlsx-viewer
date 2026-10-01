import { describe, expect, it } from 'vitest';
import {
	CURRENCY_SYMBOL_PRESETS,
	FRACTION_FORMAT_PRESETS,
	SPECIAL_FORMAT_PRESETS,
} from 'xlsx-viewer-core';

import { en } from '../../../../i18n';
import {
	CATEGORY_DESCRIPTIONS,
	CURRENCY_LABEL_KEYS,
	NUMBER_CATEGORIES,
	PRESET_LABEL_KEYS,
} from './number-draft';

describe('Number tab label keys', () => {
	it('names every Fraction and Special type with the English label of the core', () => {
		for (const preset of [...FRACTION_FORMAT_PRESETS, ...SPECIAL_FORMAT_PRESETS]) {
			const key = PRESET_LABEL_KEYS[preset.code];
			expect(key, preset.code).toBeDefined();
			expect(en[key!]).toBe(preset.label);
		}
	});

	it('names every currency symbol except the bare $ with the English label of the core', () => {
		for (const preset of CURRENCY_SYMBOL_PRESETS) {
			const key = CURRENCY_LABEL_KEYS[preset.id];
			if (preset.id === '$') {
				expect(key).toBeUndefined();
			} else {
				expect(key, preset.id).toBeDefined();
				expect(en[key!]).toBe(preset.label);
			}
		}
	});

	it('lists the twelve categories, and describes all but Fraction and Scientific', () => {
		expect(NUMBER_CATEGORIES.map((c) => en[c.labelKey])).toEqual([
			'General',
			'Number',
			'Currency',
			'Accounting',
			'Date',
			'Time',
			'Percentage',
			'Fraction',
			'Scientific',
			'Text',
			'Special',
			'Custom',
		]);
		const undescribed = NUMBER_CATEGORIES.filter((c) => !CATEGORY_DESCRIPTIONS[c.id]);
		expect(undescribed.map((c) => c.id)).toEqual(['fraction', 'scientific']);
	});
});
