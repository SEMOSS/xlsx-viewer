import { describe, expect, it } from 'vitest';
import type { CfIconSet, Cfvo, IconSetName } from '../types/conditional-format';
import { defaultIconCfvos, iconForThreshold, selectIconThreshold } from './icon-set';
import { ICON_SET_GROUPS, ICON_SETS } from './icon-sets';
import { addRule, cfFixture, setColumn } from './test-utils/fixtures';

const ALL_SETS: IconSetName[] = [
	'3Arrows',
	'3ArrowsGray',
	'3Flags',
	'3TrafficLights1',
	'3TrafficLights2',
	'3Signs',
	'3Symbols',
	'3Symbols2',
	'3Stars',
	'3Triangles',
	'4Arrows',
	'4ArrowsGray',
	'4RedToBlack',
	'4Rating',
	'4TrafficLights',
	'5Arrows',
	'5ArrowsGray',
	'5Rating',
	'5Quarters',
	'5Boxes',
	'NoIcons',
];

describe('ICON_SETS', () => {
	it('describes every icon set with the right number of icons', () => {
		for (const name of ALL_SETS) {
			const set = ICON_SETS[name];
			expect(set.name).toBe(name);
			const expected = name === 'NoIcons' ? 0 : Number(name[0]);
			expect(set.count).toBe(expected);
			expect(set.icons).toHaveLength(expected);
			for (const icon of set.icons) {
				expect(icon.color).toMatch(/^#[0-9A-F]{6}$/);
			}
		}
		expect(Object.keys(ICON_SETS).sort()).toEqual([...ALL_SETS].sort());
	});

	it('orders icons from the lowest values to the highest', () => {
		expect(ICON_SETS['3Arrows'].icons.map((i) => i.shape)).toEqual([
			'arrowDown',
			'arrowRight',
			'arrowUp',
		]);
		expect(ICON_SETS['5Arrows'].icons.map((i) => i.shape)).toEqual([
			'arrowDown',
			'arrowDownRight',
			'arrowRight',
			'arrowUpRight',
			'arrowUp',
		]);
		expect(ICON_SETS['3TrafficLights1'].icons.map((i) => i.color)).toEqual([
			'#E0393E',
			'#FFC000',
			'#00B050',
		]);
		expect(ICON_SETS['3Stars'].icons.map((i) => i.shape)).toEqual([
			'starEmpty',
			'starHalf',
			'star',
		]);
		expect(ICON_SETS['5Quarters'].icons.map((i) => i.fill)).toEqual([0, 0.25, 0.5, 0.75, 1]);
		expect(ICON_SETS['4Rating'].icons.map((i) => i.fill)).toEqual([0.25, 0.5, 0.75, 1]);
		expect(ICON_SETS['3Symbols'].icons.every((i) => i.circled)).toBe(true);
		expect(ICON_SETS['3Symbols2'].icons.some((i) => i.circled)).toBe(false);
		expect(new Set(ICON_SETS['3ArrowsGray'].icons.map((i) => i.color))).toEqual(
			new Set(['#808080']),
		);
		expect(ICON_SETS['3TrafficLights2'].icons[0].shape).toBe('trafficLight');
	});

	it('groups the gallery like Excel', () => {
		const grouped = ICON_SET_GROUPS.flatMap((g) => g.sets);
		expect(grouped).toHaveLength(20);
		expect(new Set(grouped).size).toBe(20);
		expect(grouped).not.toContain('NoIcons');
	});
});

describe('icon selection', () => {
	it('builds Excel default thresholds', () => {
		expect(defaultIconCfvos(3).map((c) => c.value)).toEqual(['0', '33', '67']);
		expect(defaultIconCfvos(4).map((c) => c.value)).toEqual(['0', '25', '50', '75']);
		expect(defaultIconCfvos(5).map((c) => c.value)).toEqual(['0', '20', '40', '60', '80']);
	});

	it('checks thresholds from the top with >= or >', () => {
		expect(selectIconThreshold(67, [0, 33, 67], [true, true, true])).toBe(2);
		expect(selectIconThreshold(67, [0, 33, 67], [true, true, false])).toBe(1);
		expect(selectIconThreshold(-100, [0, 33, 67], [true, true, true])).toBe(0);
		expect(selectIconThreshold(50, [0, 60, 40], [true, true, true])).toBe(2);
	});

	it('maps thresholds to icons with reverse and custom icons', () => {
		const set: CfIconSet = { iconSet: '3Arrows', cfvos: [] };
		expect(iconForThreshold(set, 2, 3)).toEqual({ iconSet: '3Arrows', index: 2 });
		expect(iconForThreshold({ ...set, reverse: true }, 2, 3)).toEqual({
			iconSet: '3Arrows',
			index: 0,
		});
		const custom: CfIconSet = {
			...set,
			reverse: true,
			customIcons: [
				{ iconSet: '3Flags', iconId: 1 },
				{ iconSet: 'NoIcons', iconId: 0 },
				{ iconSet: '5Rating', iconId: 9 },
			],
		};
		expect(iconForThreshold(custom, 0, 3)).toEqual({ iconSet: '3Flags', index: 1 });
		expect(iconForThreshold(custom, 1, 3)).toBeNull();
		expect(iconForThreshold(custom, 2, 3)).toBeNull();
	});
});

describe('icon set rules', () => {
	function fixture(
		iconSet: Partial<CfIconSet> & { iconSet: IconSetName },
		values: Array<number | string> = [0, 10, 20, 30, 40, 50, 60, 70, 80, 90, 100],
	) {
		const fx = cfFixture({ Z1: 50 });
		setColumn(fx.sheet, 'A1', values);
		addRule(fx.sheet, `A1:A${values.length}`, {
			type: 'iconSet',
			iconSet: { cfvos: [], ...iconSet },
		});
		const index = (row: number) => fx.at(`A${row}`)?.icon?.index;
		return { fx, index };
	}
	const pct = (...values: string[]): Cfvo[] => values.map((value) => ({ type: 'percent', value }));

	it('uses percent thresholds of the value range', () => {
		const { fx, index } = fixture({ iconSet: '3Arrows', cfvos: pct('0', '33', '67') });
		expect([1, 4, 5, 7, 8, 11].map(index)).toEqual([0, 0, 1, 1, 2, 2]);
		expect(fx.at('A1')?.icon?.iconSet).toBe('3Arrows');
	});

	it('falls back to default thresholds when cfvos are missing', () => {
		const { index } = fixture({ iconSet: '5Rating' });
		expect([1, 3, 5, 7, 9, 11].map(index)).toEqual([0, 1, 2, 3, 4, 4]);
	});

	it('supports num, percentile and formula thresholds with gte flags', () => {
		const num = fixture({
			iconSet: '3Flags',
			cfvos: [
				{ type: 'num', value: '0' },
				{ type: 'num', value: '50', gte: false },
				{ type: 'num', value: '80' },
			],
		});
		expect([5, 6, 7, 9].map(num.index)).toEqual([0, 0, 1, 2]);
		const pctl = fixture({
			iconSet: '3Flags',
			cfvos: [
				{ type: 'percentile', value: '0' },
				{ type: 'percentile', value: '50' },
				{ type: 'percentile', value: '90' },
			],
		});
		expect([5, 6, 10, 11].map(pctl.index)).toEqual([0, 1, 2, 2]);
		const formula = fixture({
			iconSet: '3Flags',
			cfvos: [
				{ type: 'num', value: '0' },
				{ type: 'formula', value: '$Z$1' },
				{ type: 'formula', value: '$Z$1+40' },
			],
		});
		expect([5, 6, 10].map(formula.index)).toEqual([0, 1, 2]);
	});

	it('reverses icon order', () => {
		const { index } = fixture({ iconSet: '3Arrows', cfvos: pct('0', '33', '67'), reverse: true });
		expect([1, 11].map(index)).toEqual([2, 0]);
	});

	it('hides the value for "show icon only"', () => {
		const { fx } = fixture({ iconSet: '3Arrows', cfvos: pct('0', '33', '67'), showValue: false });
		expect(fx.at('A1')?.hideValue).toBe(true);
	});

	it('supports custom icons including "No Cell Icon"', () => {
		const { fx } = fixture({
			iconSet: '3Arrows',
			cfvos: pct('0', '33', '67'),
			customIcons: [
				{ iconSet: 'NoIcons', iconId: 0 },
				{ iconSet: '3Stars', iconId: 1 },
				{ iconSet: '3TrafficLights1', iconId: 2 },
			],
		});
		expect(fx.at('A1')).toBeUndefined();
		expect(fx.at('A6')?.icon).toEqual({ iconSet: '3Stars', index: 1 });
		expect(fx.at('A11')?.icon).toEqual({ iconSet: '3TrafficLights1', index: 2 });
	});

	it('handles 4 and 5 icon sets and ignores text', () => {
		const four = fixture({ iconSet: '4TrafficLights' }, [0, 25, 50, 75, 100, 'n/a']);
		expect([1, 2, 3, 4, 5].map(four.index)).toEqual([0, 1, 2, 3, 3]);
		expect(four.fx.at('A6')).toBeUndefined();
	});

	it('hides the value even when the matching custom icon is "No Cell Icon"', () => {
		const { fx } = fixture({
			iconSet: '3Arrows',
			cfvos: pct('0', '33', '67'),
			showValue: false,
			customIcons: [
				{ iconSet: 'NoIcons', iconId: 0 },
				{ iconSet: '3Arrows', iconId: 1 },
				{ iconSet: '3Arrows', iconId: 2 },
			],
		});
		expect(fx.at('A1')).toEqual({ hideValue: true });
		expect(fx.at('A11')).toEqual({ icon: { iconSet: '3Arrows', index: 2 }, hideValue: true });
	});

	it('lets only the highest-priority icon set decide', () => {
		const { fx } = fixture({ iconSet: '3Arrows', cfvos: pct('0', '33', '67') });
		addRule(fx.sheet, 'A1:A11', {
			type: 'iconSet',
			priority: 1_000_000,
			iconSet: { iconSet: '3Flags', cfvos: pct('0', '33', '67') },
		});
		fx.engine.invalidate();
		expect(fx.at('A11')?.icon?.iconSet).toBe('3Arrows');
	});
});

describe('color scale rules', () => {
	function fixture(
		cfvos: Cfvo[],
		colors: string[],
		values: Array<number | string | null> = [0, 25, 50, 75, 100],
	) {
		const fx = cfFixture({ Z1: 100 });
		setColumn(fx.sheet, 'A1', values);
		addRule(fx.sheet, `A1:A${values.length}`, {
			type: 'colorScale',
			colorScale: { cfvos, colors: colors.map((rgb) => ({ rgb })) },
		});
		return fx;
	}

	it('interpolates a 3-color min / percentile / max scale', () => {
		const fx = fixture(
			[{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
			['FFF8696B', 'FFFFEB84', 'FF63BE7B'],
		);
		expect(fx.at('A1')?.colorScale).toBe('#F8696B');
		expect(fx.at('A3')?.colorScale).toBe('#FFEB84');
		expect(fx.at('A5')?.colorScale).toBe('#63BE7B');
		expect(fx.at('A2')?.colorScale).toBe('#FCAA78');
	});

	it('interpolates a 2-color scale with num and formula thresholds', () => {
		const fx = fixture(
			[
				{ type: 'num', value: '0' },
				{ type: 'formula', value: '$Z$1*2' },
			],
			['FFFFFFFF', 'FF000000'],
		);
		expect(fx.at('A5')?.colorScale).toBe('#808080');
	});

	it('supports percent thresholds and theme colors', () => {
		const fx = cfFixture();
		setColumn(fx.sheet, 'A1', [10, 20]);
		addRule(fx.sheet, 'A1:A2', {
			type: 'colorScale',
			colorScale: {
				cfvos: [
					{ type: 'percent', value: '0' },
					{ type: 'percent', value: '100' },
				],
				colors: [{ theme: 0 }, { theme: 4 }],
			},
		});
		expect(fx.at('A1')?.colorScale).toBe('#FFFFFF');
		expect(fx.at('A2')?.colorScale).toBe('#4472C4');
	});

	it('ignores non-numeric cells and all-equal ranges use the first color', () => {
		const fx = fixture(
			[{ type: 'min' }, { type: 'max' }],
			['FFFF0000', 'FF00FF00'],
			[5, 5, 'x', null],
		);
		expect(fx.at('A1')?.colorScale).toBe('#FF0000');
		expect(fx.at('A3')).toBeUndefined();
		expect(fx.at('A4')).toBeUndefined();
	});
});
