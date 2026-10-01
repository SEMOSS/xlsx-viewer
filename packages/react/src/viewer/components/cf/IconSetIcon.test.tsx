import { render } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import { ICON_SETS, type IconSetName } from 'xlsx-viewer-core';
import { IconSetIcon } from './IconSetIcon';

const SETS = Object.keys(ICON_SETS) as IconSetName[];

function svgOf(iconSet: IconSetName, index: number, size?: number) {
	const { container } = render(<IconSetIcon iconSet={iconSet} index={index} size={size} />);
	return container.querySelector('svg');
}

describe('IconSetIcon', () => {
	it('renders every icon of every set as an SVG in the descriptor color (empty gauges are gray)', () => {
		let rendered = 0;
		for (const name of SETS) {
			ICON_SETS[name].icons.forEach((icon, index) => {
				const svg = svgOf(name, index);
				expect(svg, `${name}[${index}]`).not.toBeNull();
				expect(svg!.getAttribute('viewBox')).toBe('0 0 16 16');
				expect(svg!.getAttribute('data-shape')).toBe(icon.shape);
				expect(svg!.childElementCount).toBeGreaterThan(0);
				const emptyGauge = (icon.shape === 'rating' || icon.shape === 'box') && icon.fill === 0;
				if (emptyGauge) {
					expect(svg!.innerHTML).not.toContain(icon.color);
				} else {
					expect(svg!.innerHTML).toContain(icon.color);
				}
				rendered++;
			});
		}
		expect(rendered).toBe(3 * 10 + 4 * 5 + 5 * 5);
	});

	it('labels the icon for assistive technology', () => {
		const svg = svgOf('3Arrows', 2);
		expect(svg!.getAttribute('role')).toBe('img');
		expect(svg!.getAttribute('aria-label')).toBe('3 Arrows (Colored), icon 3 of 3');
		expect(svg!.getAttribute('data-icon-set')).toBe('3Arrows');
		expect(svg!.getAttribute('data-icon-index')).toBe('2');
	});

	it('defaults to 16px and honours the size prop', () => {
		expect(svgOf('3Flags', 0)!.getAttribute('width')).toBe('16');
		const big = svgOf('3Flags', 0, 24)!;
		expect(big.getAttribute('width')).toBe('24');
		expect(big.getAttribute('height')).toBe('24');
	});

	it('renders nothing for unknown icons', () => {
		expect(svgOf('3Arrows', 3)).toBeNull();
		expect(svgOf('3Arrows', -1)).toBeNull();
		expect(svgOf('NoIcons', 0)).toBeNull();
	});

	it('draws shape details', () => {
		expect(svgOf('3Arrows', 2)!.querySelector('polygon')).not.toBeNull();
		expect(svgOf('4Arrows', 2)!.querySelector('polygon')!.getAttribute('transform')).toBe(
			'rotate(45 8 8)',
		);
		expect(svgOf('4Arrows', 1)!.querySelector('polygon')!.getAttribute('transform')).toBe(
			'rotate(135 8 8)',
		);
		expect(svgOf('3TrafficLights2', 0)!.querySelectorAll('circle')).toHaveLength(2);
		expect(svgOf('3Symbols', 2)!.querySelector('circle')!.getAttribute('fill')).toBe(
			ICON_SETS['3Symbols'].icons[2].color,
		);
		expect(svgOf('3Symbols', 2)!.querySelector('path')!.getAttribute('stroke')).toBe('#FFFFFF');
		expect(svgOf('3Symbols2', 0)!.querySelector('circle')).toBeNull();
		expect(svgOf('3Stars', 1)!.querySelectorAll('polygon')).toHaveLength(2);
		expect(svgOf('3Stars', 0)!.querySelector('polygon')!.getAttribute('fill')).toBe('#FFFFFF');
		expect(svgOf('3Flags', 1)!.querySelector('path')).not.toBeNull();
	});

	it('fills quarters, ratings and boxes progressively', () => {
		expect(svgOf('5Quarters', 0)!.querySelector('path')).toBeNull();
		expect(svgOf('5Quarters', 2)!.querySelector('path')!.getAttribute('d')).toContain(
			'A6.5 6.5 0 0 1 8 14.5',
		);
		expect(svgOf('5Quarters', 4)!.querySelector('circle')!.getAttribute('fill')).toBe(
			ICON_SETS['5Quarters'].icons[4].color,
		);
		const lit = (name: IconSetName, index: number) =>
			[...svgOf(name, index)!.querySelectorAll('rect')].filter(
				(r) => r.getAttribute('fill') !== '#D9D9D9',
			).length;
		expect([0, 1, 2, 3, 4].map((i) => lit('5Rating', i))).toEqual([0, 1, 2, 3, 4]);
		expect([0, 1, 2, 3].map((i) => lit('4Rating', i))).toEqual([1, 2, 3, 4]);
		expect([0, 1, 2, 3, 4].map((i) => lit('5Boxes', i))).toEqual([0, 1, 2, 3, 4]);
	});
});
