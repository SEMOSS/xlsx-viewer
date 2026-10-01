import { describe, expect, it } from 'vitest';

import {
	attrBool,
	attrNum,
	child,
	childFlag,
	childVal,
	childrenLocal,
	cloneElement,
	el,
	findAllDescendantsLocal,
	localName,
	namePrefix,
	textContent,
} from './dom';
import {
	decodeOoxmlEscapes,
	decodeXmlEntities,
	encodeOoxmlEscapes,
	escapeXmlAttr,
	escapeXmlText,
} from './escape';
import { parseXml } from './parser';
import { serializeXml } from './serializer';

const doc = parseXml(
	'<worksheet xmlns:x14="urn:x14"><sheetPr><tabColor rgb="FF00FF00"/></sheetPr><font><b/><i val="0"/><sz val="11"/></font>' +
		'<x14:dataBar minLength="10" showValue="false"/><c r="A1"><v>1</v></c><c r="B1"><is><t xml:space="preserve"> two </t></is></c></worksheet>',
).root;

describe('XML DOM helpers', () => {
	it('splits qualified names', () => {
		expect(localName('x14:dataBar')).toBe('dataBar');
		expect(namePrefix('x14:dataBar')).toBe('x14');
		expect(localName('c')).toBe('c');
	});

	it('reads typed attributes and toggle children', () => {
		const font = child(doc, 'font');
		expect(childFlag(font, 'b')).toBe(true);
		expect(childFlag(font, 'i')).toBe(false);
		expect(childFlag(font, 'u')).toBeUndefined();
		expect(childVal(font, 'sz')).toBe('11');
		const bar = child(doc, 'x14:dataBar');
		expect(attrNum(bar, 'minLength')).toBe(10);
		expect(attrNum(bar, 'missing', 5)).toBe(5);
		expect(attrBool(bar, 'showValue')).toBe(false);
		expect(attrBool(bar, 'missing', true)).toBe(true);
	});

	it('finds children and descendants by local name and reads text', () => {
		expect(childrenLocal(doc, 'c')).toHaveLength(2);
		expect(findAllDescendantsLocal(doc, 't').map(textContent)).toEqual([' two ']);
	});

	it('builds and clones elements without sharing children', () => {
		const built = el('row', { r: 3, spans: '1:2' }, [el('c', { r: 'A3' })]);
		expect(serializeXml(built)).toContain('<row r="3" spans="1:2"><c r="A3"/></row>');
		const copy = cloneElement(built);
		(copy.children[0] as typeof built).attrs.r = 'B3';
		expect(serializeXml(built)).toContain('A3');
	});
});

describe('XML escapes', () => {
	it('escapes text and attributes and decodes entities', () => {
		expect(escapeXmlText('a < b & c > d')).toBe('a &lt; b &amp; c &gt; d');
		expect(escapeXmlAttr('say "hi" & \'bye\'')).toContain('&quot;');
		expect(decodeXmlEntities('&lt;&#65;&#x42;&amp;&quot;')).toBe('<AB&"');
	});

	it('round-trips OOXML _xHHHH_ escapes for control characters', () => {
		expect(decodeOoxmlEscapes('tab_x0009_here')).toBe('tab\there');
		const literal = 'keep _x0041_ literally';
		expect(decodeOoxmlEscapes(encodeOoxmlEscapes(literal))).toBe(literal);
	});
});
