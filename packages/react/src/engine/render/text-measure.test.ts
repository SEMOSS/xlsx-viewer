import { afterEach, describe, expect, it } from 'vitest';

import { measureLines, measureText, setMeasureContextForTests } from './text-measure';

interface FakeContext {
	font: string;
	calls: string[];
	measureText(text: string): { width: number };
}

/** A canvas stand-in whose glyphs are all `perChar` pixels wide. */
function fakeContext(perChar: number): FakeContext {
	const fake: FakeContext = {
		font: '',
		calls: [],
		measureText(text: string) {
			fake.calls.push(`${fake.font}|${text}`);
			return { width: text.length * perChar };
		},
	};
	return fake;
}

function install(fake: FakeContext): FakeContext {
	setMeasureContextForTests(fake as unknown as CanvasRenderingContext2D);
	return fake;
}

afterEach(() => {
	setMeasureContextForTests(null);
});

describe('measureText without a canvas (estimate)', () => {
	it('weights narrow, wide, capital and other characters', () => {
		expect(measureText('i', '10px x')).toBeCloseTo(3, 6);
		expect(measureText('|', '10px x')).toBeCloseTo(3, 6);
		expect(measureText('m', '10px x')).toBeCloseTo(8.5, 6);
		expect(measureText('W', '10px x')).toBeCloseTo(8.5, 6);
		expect(measureText('A', '10px x')).toBeCloseTo(6.2, 6);
		expect(measureText('7', '10px x')).toBeCloseTo(6.2, 6);
		expect(measureText('a', '10px x')).toBeCloseTo(5.2, 6);
		expect(measureText(' ', '10px x')).toBeCloseTo(5.2, 6);
		expect(measureText('中', '10px x')).toBeCloseTo(10, 6);
		expect(measureText('Hello', '10px x')).toBeCloseTo(22.6, 6);
	});

	it('scales with the font size of the canvas font string', () => {
		const font = "normal 400 14.666666666666666px Calibri, 'Segoe UI', sans-serif";
		expect(measureText('0', font)).toBeCloseTo(0.62 * 14.6667, 3);
		expect(measureText('a', '20px x')).toBeCloseTo(2 * measureText('a', '10px x'), 6);
		// Without a pixel size the estimate assumes 14px.
		expect(measureText('a', 'Arial')).toBeCloseTo(7.28, 6);
	});

	it('makes bold text 6% wider', () => {
		expect(measureText('abc', 'normal 700 10px x')).toBeCloseTo(15.6 * 1.06, 6);
		expect(measureText('abc', 'bold 10px x')).toBeCloseTo(15.6 * 1.06, 6);
		expect(measureText('abc', 'normal 400 10px x')).toBeCloseTo(15.6, 6);
	});

	it('measures empty text as zero', () => {
		expect(measureText('', '10px x')).toBe(0);
	});
});

describe('measureText with a canvas', () => {
	it('measures with the canvas in the requested font', () => {
		const fake = install(fakeContext(10));
		expect(measureText('abc', 'italic 400 12px Arial')).toBe(30);
		expect(fake.font).toBe('italic 400 12px Arial');
		expect(fake.calls).toEqual(['italic 400 12px Arial|abc']);
	});

	it('does not call the canvas for empty text', () => {
		const fake = install(fakeContext(10));
		expect(measureText('', '12px Arial')).toBe(0);
		expect(fake.calls).toEqual([]);
	});

	it('caches widths per font and text', () => {
		const fake = install(fakeContext(10));
		measureText('abc', '12px A');
		measureText('abc', '12px A');
		expect(fake.calls).toHaveLength(1);
		measureText('abc', '12px B');
		measureText('abcd', '12px A');
		expect(fake.calls).toHaveLength(3);
	});

	it('drops cached widths when the measuring context is replaced', () => {
		install(fakeContext(10));
		expect(measureText('abc', '12px A')).toBe(30);
		install(fakeContext(20));
		expect(measureText('abc', '12px A')).toBe(60);
	});

	it('bounds the cache by clearing it once it holds more than 20000 entries', () => {
		const fake = install(fakeContext(1));
		for (let i = 0; i < 20002; i++) {
			measureText(`k${i}`, '1px x');
		}
		expect(fake.calls).toHaveLength(20002);
		// The last entry was stored after the clear; the first one was dropped.
		measureText('k20001', '1px x');
		expect(fake.calls).toHaveLength(20002);
		measureText('k0', '1px x');
		expect(fake.calls).toHaveLength(20003);
	});
});

describe('measureLines', () => {
	it('returns the widest line', () => {
		install(fakeContext(10));
		expect(measureLines('ab\nabcd\nabc', '12px A')).toBe(40);
		expect(measureLines('abc', '12px A')).toBe(measureText('abc', '12px A'));
		expect(measureLines('', '12px A')).toBe(0);
		expect(measureLines('\n\n', '12px A')).toBe(0);
	});
});
