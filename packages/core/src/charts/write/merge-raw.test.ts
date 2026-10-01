import { describe, expect, it } from 'vitest';
import { el } from '../../xml/dom';
import { parseXmlRoot } from '../../xml/parser';
import { serializeElement } from '../../xml/serializer';
import { mergeRawChildren, rawChildByKey } from './merge-raw';

const raw = parseXmlRoot(
	'<c:catAx xmlns:c="c"><c:axId val="7"/><c:scaling/><c:axPos val="b"/><c:crossAx val="8"/><c:auto val="0"/><c:lblOffset val="50"/><c:noMultiLvlLbl val="1"/><c:extLst><c:ext uri="x"/></c:extLst></c:catAx>',
);

describe('mergeRawChildren', () => {
	it('inserts unowned originals in schema order and prefers original defaults', () => {
		const generated = el('c:catAx', {}, [
			el('c:axId', { val: 7 }),
			el('c:scaling'),
			el('c:axPos', { val: 'l' }),
			el('c:crossAx', { val: 8 }),
			el('c:auto', { val: 1 }),
			el('c:noMultiLvlLbl', { val: 0 }),
		]);
		mergeRawChildren(generated, raw, {
			order: 'catAx',
			owned: ['axId', 'scaling', 'axPos', 'crossAx', 'noMultiLvlLbl'],
			preferRaw: ['auto'],
		});
		expect(serializeElement(generated)).toBe(
			'<c:catAx><c:axId val="7"/><c:scaling/><c:axPos val="l"/><c:crossAx val="8"/><c:auto val="0"/><c:lblOffset val="50"/><c:noMultiLvlLbl val="0"/><c:extLst><c:ext uri="x"/></c:extLst></c:catAx>',
		);
	});

	it('does not re-add owned elements the model removed', () => {
		const generated = el('c:catAx', {}, [el('c:axId', { val: 7 })]);
		mergeRawChildren(generated, raw, {
			order: 'catAx',
			owned: [
				'axId',
				'scaling',
				'axPos',
				'crossAx',
				'auto',
				'lblOffset',
				'noMultiLvlLbl',
				'extLst',
			],
		});
		expect(serializeElement(generated)).toBe('<c:catAx><c:axId val="7"/></c:catAx>');
	});

	it('finds original children by key', () => {
		const plot = parseXmlRoot(
			'<c:barChart xmlns:c="c"><c:ser><c:idx val="0"/></c:ser><c:ser><c:idx val="3"/></c:ser></c:barChart>',
		);
		expect(rawChildByKey(plot, 'ser', 'idx', 3)).toBe(plot.children[1]);
		expect(rawChildByKey(plot, 'ser', 'idx', 9)).toBeUndefined();
		expect(rawChildByKey(undefined, 'ser', 'idx', 0)).toBeUndefined();
	});
});
