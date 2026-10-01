import { describe, expect, it } from 'vitest';

import { parseSharedStrings } from './shared-strings';
import { MAIN_NS } from './test-utils/package-builder';

function sst(items: string, prefix = ''): string {
	const p = prefix ? `${prefix}:` : '';
	const ns = prefix ? `xmlns:${prefix}="${MAIN_NS}"` : `xmlns="${MAIN_NS}"`;
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${p}sst ${ns} count="9" uniqueCount="9">${items}</${p}sst>`;
}

describe('parseSharedStrings', () => {
	it('reads plain strings, entities, escapes and whitespace', () => {
		const table = parseSharedStrings(
			sst(
				'<si><t>Hello</t></si>' +
					'<si><t xml:space="preserve">  padded  </t></si>' +
					'<si><t>a &amp; b &lt;c&gt; &#x41;</t></si>' +
					'<si><t>line_x000a_break and _x005F_x0041_ literal</t></si>' +
					'<si><t/></si>' +
					'<si/>' +
					'<si>\n  <t>indented</t>\n</si>' +
					'<si><t><![CDATA[<cdata>]]></t></si>',
			),
		);
		expect(table.strings).toEqual([
			'Hello',
			'  padded  ',
			'a & b <c> A',
			'line\nbreak and _x0041_ literal',
			'',
			'',
			'indented',
			'<cdata>',
		]);
		expect(table.rich.size).toBe(0);
	});

	it('reads rich runs with fonts and drops phonetic runs', () => {
		const table = parseSharedStrings(
			sst(
				'<si><r><rPr><b/><sz val="11"/><color theme="1"/><rFont val="Calibri"/><family val="2"/><scheme val="minor"/></rPr><t>Bold</t></r>' +
					'<r><t xml:space="preserve"> and </t></r>' +
					'<r><rPr><i/><u val="double"/><vertAlign val="subscript"/></rPr><t>more</t></r></si>' +
					'<si><t>東京</t><rPh sb="0" eb="2"><t>トウキョウ</t></rPh><phoneticPr fontId="1"/></si>' +
					'<si><r><t>unformatted</t></r><r><t> runs</t></r></si>',
			),
		);
		expect(table.strings).toEqual(['Bold and more', '東京', 'unformatted runs']);
		expect(table.rich.get(0)).toEqual([
			{
				text: 'Bold',
				font: {
					bold: true,
					size: 11,
					color: { theme: 1 },
					name: 'Calibri',
					family: 2,
					scheme: 'minor',
				},
			},
			{ text: ' and ' },
			{ text: 'more', font: { italic: true, underline: 'double', vertAlign: 'subscript' } },
		]);
		expect(table.rich.has(1)).toBe(false);
		expect(table.rich.has(2)).toBe(false);
	});

	it('reads a table whose elements use a namespace prefix', () => {
		const table = parseSharedStrings(
			sst(
				'<x:si><x:t>one</x:t></x:si><x:si><x:r><x:rPr><x:b/></x:rPr><x:t>two</x:t></x:r></x:si>',
				'x',
			),
		);
		expect(table.strings).toEqual(['one', 'two']);
		expect(table.rich.get(1)).toEqual([{ text: 'two', font: { bold: true } }]);
	});

	it('returns an empty table for an empty part', () => {
		expect(parseSharedStrings(sst('')).strings).toEqual([]);
		expect(parseSharedStrings('').strings).toEqual([]);
	});
});
