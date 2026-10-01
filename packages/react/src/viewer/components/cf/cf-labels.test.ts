import {
	COLOR_SCALE_PRESETS,
	DATA_BAR_PRESETS,
	ICON_SET_GROUPS,
	ICON_SETS,
	type IconSetName,
} from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { createTranslator } from '../../../i18n';
import {
	COLOR_SCALE_KEYS,
	DATA_BAR_KEYS,
	ICON_SET_GROUP_KEYS,
	ICON_SET_KEYS,
	presetLabel,
} from './cf-labels';

const t = createTranslator();

describe('conditional formatting gallery labels', () => {
	it('has a key for every core preset, worded like the core in English', () => {
		for (const { name } of DATA_BAR_PRESETS) {
			expect(DATA_BAR_KEYS[name]).toBeDefined();
			expect(presetLabel(DATA_BAR_KEYS, name, t)).toBe(name);
		}
		for (const { name } of COLOR_SCALE_PRESETS) {
			expect(COLOR_SCALE_KEYS[name]).toBeDefined();
			expect(presetLabel(COLOR_SCALE_KEYS, name, t)).toBe(name);
		}
		for (const { label } of ICON_SET_GROUPS) {
			expect(ICON_SET_GROUP_KEYS[label]).toBeDefined();
			expect(presetLabel(ICON_SET_GROUP_KEYS, label, t)).toBe(label);
		}
		for (const name of Object.keys(ICON_SETS) as IconSetName[]) {
			expect(t(ICON_SET_KEYS[name])).toBe(ICON_SETS[name].label);
		}
	});

	it('translates through the active dictionary', () => {
		const de = createTranslator({ 'cf.dataBar.blue': 'Blaue Datenleiste' });
		expect(presetLabel(DATA_BAR_KEYS, 'Blue Data Bar', de)).toBe('Blaue Datenleiste');
	});

	it('keeps the core name of a preset without a key', () => {
		expect(presetLabel(DATA_BAR_KEYS, 'Teal Data Bar', t)).toBe('Teal Data Bar');
	});
});
