import { describe, expect, it } from 'vitest';

import { ERRORS } from '../../../values/errors';
import { arr, call, MockContext } from '../../test-utils/mock-context';
import { bahtText } from './bahttext';
import { MISC_FUNCTIONS } from './index';

const fn = (name: string) => MISC_FUNCTIONS.find((f) => f.name === name)!;

describe('PERCENTOF', () => {
	it('divides the subset sum by the total', () => {
		expect(call(fn('PERCENTOF'), [arr([1, 3]), arr([1, 3, 4, 2])])).toBe(0.4);
		expect(call(fn('PERCENTOF'), [arr([1]), arr([0])])).toBe(ERRORS.DIV0);
	});
});

describe('ENCODEURL', () => {
	it('percent-encodes everything but unreserved characters', () => {
		expect(call(fn('ENCODEURL'), ['a b/c?d=é&e*'])).toBe('a%20b%2Fc%3Fd%3D%C3%A9%26e%2A');
		expect(call(fn('ENCODEURL'), ['safe-_.~'])).toBe('safe-_.~');
	});
});

describe('BAHTTEXT', () => {
	it('spells amounts in Thai baht and satang', () => {
		expect(bahtText(1234)).toBe('หนึ่งพันสองร้อยสามสิบสี่บาทถ้วน');
		expect(bahtText(21)).toBe('ยี่สิบเอ็ดบาทถ้วน');
		expect(bahtText(11.25)).toBe('สิบเอ็ดบาทยี่สิบห้าสตางค์');
		expect(bahtText(0.5)).toBe('ห้าสิบสตางค์');
		expect(bahtText(0)).toBe('ศูนย์บาทถ้วน');
		expect(bahtText(-1)).toBe('ลบหนึ่งบาทถ้วน');
		expect(bahtText(2_000_001)).toBe('สองล้านหนึ่งบาทถ้วน');
	});
});

describe('TRIMRANGE', () => {
	it('drops blank outer rows and columns of a reference', () => {
		const ctx = new MockContext({ B3: 1, C5: 'x' });
		const ref = ctx.ref('A1:E9');
		expect(call(fn('TRIMRANGE'), [ref], ctx)).toEqual({
			kind: 'ref',
			sheet: 0,
			range: { top: 2, left: 1, bottom: 4, right: 2 },
		});
		expect(call(fn('TRIMRANGE'), [ref, 2, 0], ctx)).toEqual({
			kind: 'ref',
			sheet: 0,
			range: { top: 0, left: 0, bottom: 4, right: 4 },
		});
		expect(call(fn('TRIMRANGE'), [ctx.ref('G1:H2')], ctx)).toBe(ERRORS.REF);
	});
});

describe('online-only functions', () => {
	it('report that they are unavailable', () => {
		expect(call(fn('WEBSERVICE'), ['https://example.com'])).toBe(ERRORS.BLOCKED);
		expect(call(fn('CUBEVALUE'), ['conn'])).toBe(ERRORS.NA);
	});
});

describe('FILTERXML', () => {
	const xml = '<t><s>alpha</s><s>12</s><s></s><s a="x">beta</s></t>';
	it('splits text with the classic //s pattern', () => {
		expect(call(fn('FILTERXML'), [xml, '//s'])).toEqual(arr(['alpha'], [12], [''], ['beta']));
		expect(call(fn('FILTERXML'), [xml, '//s[2]'])).toBe(12);
		expect(call(fn('FILTERXML'), [xml, '//s[last()]'])).toBe('beta');
		expect(call(fn('FILTERXML'), [xml, "//s[. != '']"])).toEqual(arr(['alpha'], [12], ['beta']));
		expect(call(fn('FILTERXML'), [xml, '//s[position() > 2]'])).toEqual(arr([''], ['beta']));
		expect(call(fn('FILTERXML'), [xml, '//s/@a'])).toBe('x');
		expect(call(fn('FILTERXML'), [xml, "//s[contains(., 'ph')]"])).toBe('alpha');
		expect(call(fn('FILTERXML'), [xml, '/t/s[position() mod 2 = 1]'])).toEqual(
			arr(['alpha'], ['']),
		);
	});

	it('is #VALUE! for no match or invalid input', () => {
		expect(call(fn('FILTERXML'), [xml, '//missing'])).toBe(ERRORS.VALUE);
		expect(call(fn('FILTERXML'), ['<broken', '//s'])).toBe(ERRORS.VALUE);
	});
});
