import { render } from '@testing-library/react';
import { describe, expect, it } from 'vitest';
import type { DataBarResult } from 'xlsx-viewer-core';
import { DataBar } from './DataBar';

const base: DataBarResult = {
	start: 0,
	end: 0.5,
	negative: false,
	color: '#638EC6',
	gradient: false,
	direction: 'ltr',
};

function draw(bar: Partial<DataBarResult>, height = 20) {
	const { container } = render(<DataBar bar={{ ...base, ...bar }} height={height} />);
	const svg = container.querySelector('svg')!;
	return {
		svg,
		rect: svg.querySelector('rect'),
		line: svg.querySelector('line'),
		gradient: svg.querySelector('linearGradient'),
	};
}

describe('DataBar', () => {
	it('positions a solid bar inside the cell box', () => {
		const { svg, rect, line, gradient } = draw({ start: 0.25, end: 0.75 });
		expect(svg.getAttribute('aria-hidden')).toBe('true');
		expect(svg.style.position).toBe('absolute');
		expect(svg.getAttribute('height')).toBe('20');
		expect(rect!.getAttribute('x')).toBe('25%');
		expect(rect!.getAttribute('width')).toBe('50%');
		expect(rect!.getAttribute('y')).toBe('2');
		expect(rect!.getAttribute('height')).toBe('16');
		expect(rect!.getAttribute('fill')).toBe('#638EC6');
		expect(rect!.getAttribute('stroke')).toBeNull();
		expect(line).toBeNull();
		expect(gradient).toBeNull();
	});

	it('fades gradients from the origin to white', () => {
		const { rect, gradient } = draw({ gradient: true });
		const id = gradient!.getAttribute('id')!;
		expect(id).toMatch(/^xv-databar-[A-Za-z0-9_-]+$/);
		expect(rect!.getAttribute('fill')).toBe(`url(#${id})`);
		expect(gradient!.getAttribute('x1')).toBe('0');
		expect(gradient!.getAttribute('x2')).toBe('1');
		const stops = gradient!.querySelectorAll('stop');
		expect(stops[0].getAttribute('stop-color')).toBe('#638EC6');
		expect(stops[1].getAttribute('stop-color')).toBe('#FFFFFF');
	});

	it('reverses the gradient for negative and right-to-left bars', () => {
		expect(draw({ gradient: true, negative: true }).gradient!.getAttribute('x1')).toBe('1');
		expect(draw({ gradient: true, direction: 'rtl' }).gradient!.getAttribute('x1')).toBe('1');
		expect(
			draw({ gradient: true, negative: true, direction: 'rtl' }).gradient!.getAttribute('x1'),
		).toBe('0');
	});

	it('gives each bar its own gradient id', () => {
		const { container } = render(
			<div>
				<DataBar bar={{ ...base, gradient: true }} height={20} />
				<DataBar bar={{ ...base, gradient: true }} height={20} />
			</div>,
		);
		const ids = [...container.querySelectorAll('linearGradient')].map((g) => g.id);
		expect(new Set(ids).size).toBe(2);
	});

	it('draws the border and the axis', () => {
		const { rect, line, svg } = draw({
			start: 0.1,
			end: 0.4,
			negative: true,
			color: '#FF0000',
			borderColor: '#800000',
			axis: 0.4,
			axisColor: '#333333',
		});
		expect(rect!.getAttribute('stroke')).toBe('#800000');
		expect(rect!.getAttribute('stroke-width')).toBe('1');
		expect(rect!.getAttribute('fill')).toBe('#FF0000');
		expect(line!.getAttribute('x1')).toBe('40%');
		expect(line!.getAttribute('y2')).toBe('20');
		expect(line!.getAttribute('stroke')).toBe('#333333');
		expect(line!.getAttribute('stroke-dasharray')).toBe('2 2');
		expect(svg.getAttribute('data-negative')).toBe('true');
	});

	it('draws only the axis for zero-length bars', () => {
		const { rect, line } = draw({ start: 0.3, end: 0.3, axis: 0.3 });
		expect(rect).toBeNull();
		expect(line).not.toBeNull();
		expect(line!.getAttribute('stroke')).toBe('#000000');
	});

	it('clamps out-of-range fractions and adapts padding to short cells', () => {
		const { rect } = draw({ start: -0.2, end: 1.4 }, 6);
		expect(rect!.getAttribute('x')).toBe('0%');
		expect(rect!.getAttribute('width')).toBe('100%');
		expect(rect!.getAttribute('y')).toBe('1');
		expect(rect!.getAttribute('height')).toBe('4');
		expect(draw({}, 2).rect!.getAttribute('height')).toBe('2');
	});
});
