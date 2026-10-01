import { describe, expect, it } from 'vitest';

import { parseXml } from '../xml';
import { MAX_CELL_TEXT, SharedStringTable } from './shared-strings';
import { pick, pickAll } from './test-utils/package-check';

const ORIGINAL =
	'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
	'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="9" uniqueCount="4">' +
	'<si><t>alpha</t></si>' +
	'<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/></rPr><t>Bold</t></r><r><t xml:space="preserve"> rest</t></r></si>' +
	'<si><t>beta</t><phoneticPr fontId="1" type="noConversion"/></si>' +
	'<si><t>line_x000D_break</t></si>' +
	'</sst>';

describe('SharedStringTable', () => {
	it('keeps original entries verbatim and in order', () => {
		const sst = SharedStringTable.fromXml(ORIGINAL);
		expect(sst.size).toBe(4);
		expect(sst.grew).toBe(false);
		const out = sst.serialize(false);
		expect(out).toContain('<si><t>alpha</t></si><si><r><rPr><b/><sz val="11"/>');
		expect(out).toContain('<si><t>beta</t><phoneticPr fontId="1" type="noConversion"/></si>');
		expect(parseXml(out).root.attrs).toEqual({
			xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
			uniqueCount: '4',
		});
	});

	it('reuses original indices for equal plain and rich strings', () => {
		const sst = SharedStringTable.fromXml(ORIGINAL);
		expect(sst.internPlain('alpha')).toBe(0);
		expect(sst.internPlain('beta')).toBe(2);
		expect(sst.internPlain('line\rbreak')).toBe(3);
		expect(
			sst.internRich([
				{
					text: 'Bold',
					font: { bold: true, size: 11, color: { theme: 1 }, name: 'Calibri', family: 2 },
				},
				{ text: ' rest' },
			]),
		).toBe(1);
		expect(sst.grew).toBe(false);
		expect(sst.refs).toBe(4);
	});

	it('appends and deduplicates new strings', () => {
		const sst = SharedStringTable.fromXml(ORIGINAL);
		expect(sst.internPlain('gamma')).toBe(4);
		expect(sst.internPlain('gamma')).toBe(4);
		expect(sst.internRich([{ text: 'Bold', font: { italic: true } }, { text: '!' }])).toBe(5);
		expect(sst.internRich([{ text: 'Bold', font: { italic: true } }, { text: '!' }])).toBe(5);
		expect(sst.internRich([{ text: 'gam' }, { text: 'ma' }])).toBe(4);
		const root = parseXml(sst.serialize(true)).root;
		expect(root.attrs.count).toBe('5');
		expect(root.attrs.uniqueCount).toBe('6');
		const si = pickAll(root, 'si');
		expect(pick(si[4], 't')?.children[0]).toBe('gamma');
		expect(pickAll(si[5], 'r').map((r) => pick(r, 't')?.children[0])).toEqual(['Bold', '!']);
		expect(pick(si[5], 'r', 'rPr', 'i')).toBeDefined();
		expect(pick(pickAll(si[5], 'r')[1], 'rPr')).toBeUndefined();
	});

	it('creates a table from scratch with exact counts', () => {
		const sst = SharedStringTable.empty();
		sst.internPlain('a');
		sst.internPlain('a');
		sst.internPlain(' b ');
		const out = sst.serialize(true);
		expect(out).toContain('count="3" uniqueCount="2"');
		expect(out).toContain('<si><t xml:space="preserve"> b </t></si>');
	});

	it('truncates strings beyond the Excel limit', () => {
		const sst = SharedStringTable.empty();
		sst.internPlain('x'.repeat(MAX_CELL_TEXT + 10));
		expect(sst.truncated).toBe(1);
		const t = pick(pick(parseXml(sst.serialize(true)).root, 'si'), 't');
		expect((t?.children[0] as string).length).toBe(MAX_CELL_TEXT);
	});

	it('keeps a prefixed original table consistent', () => {
		const sst = SharedStringTable.fromXml(
			'<x:sst xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:si><x:t>a</x:t></x:si></x:sst>',
		);
		expect(sst.internPlain('a')).toBe(0);
		sst.internPlain('b');
		const out = sst.serialize(false);
		expect(out).toContain('<x:si><x:t>a</x:t></x:si><x:si><x:t>b</x:t></x:si></x:sst>');
		expect(() => parseXml(out)).not.toThrow();
	});
});
