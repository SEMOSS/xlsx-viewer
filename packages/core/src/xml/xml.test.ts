import { describe, expect, it } from 'vitest';

import { child, el, parseXml, serializeXml, setChildInOrder, textContent } from './index';

describe('parseXml / serializeXml', () => {
	it('round-trips a worksheet fragment with prefixes, entities and preserved whitespace', () => {
		const src =
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' +
			'<sheetData><row r="1" x14ac:dyDescent="0.25"><c r="A1" t="inlineStr"><is><t xml:space="preserve"> a &amp; b &lt;c&gt; </t></is></c></row></sheetData>' +
			'</worksheet>';
		const doc = parseXml(src);
		expect(doc.root.name).toBe('worksheet');
		const t = child(child(child(child(child(doc.root, 'sheetData'), 'row'), 'c'), 'is'), 't');
		expect(textContent(t)).toBe(' a & b <c> ');
		expect(serializeXml(doc)).toBe(src);
	});

	it('drops indentation whitespace between elements but keeps whitespace-only text content', () => {
		const doc = parseXml('<a>\n  <b> </b>\n  <c>x</c>\n</a>');
		expect(doc.root.children).toHaveLength(2);
		expect(textContent(child(doc.root, 'b'))).toBe(' ');
	});

	it('handles CDATA, comments, numeric entities and single-quoted attributes', () => {
		const doc = parseXml(`<a k='v "q"'><!-- note --><![CDATA[<raw>]]>&#65;&#x42;</a>`);
		expect(doc.root.attrs.k).toBe('v "q"');
		expect(textContent(doc.root)).toBe('<raw>AB');
		expect(serializeXml(doc.root)).toContain('<!-- note -->');
	});

	it('rejects mismatched tags', () => {
		expect(() => parseXml('<a><b></a>')).toThrow(/Mismatched/);
	});

	it('inserts children in schema order', () => {
		const root = el('worksheet', {}, [el('sheetData'), el('pageMargins')]);
		setChildInOrder(root, 'mergeCells', el('mergeCells'), [
			'sheetPr',
			'dimension',
			'sheetData',
			'mergeCells',
			'pageMargins',
		]);
		expect(root.children.map((c) => (c as { name: string }).name)).toEqual([
			'sheetData',
			'mergeCells',
			'pageMargins',
		]);
	});
});
