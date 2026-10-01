import { describe, expect, it } from 'vitest';

import { readWorkbook } from './read-workbook';
import { parseVmlNoteShapes } from './vml';
import { MAIN_NS, REL_BASE, buildWorkbook, worksheet } from './test-utils/package-builder';

const TABLE_XML =
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<table xmlns="${MAIN_NS}" id="3" name="Sales" displayName="Sales_Table" ref="A1:C5" ` +
	'headerRowCount="1" totalsRowCount="1" totalsRowShown="1" headerRowDxfId="4" dataDxfId="5" totalsRowDxfId="6" comment="Quarterly">' +
	'<autoFilter ref="A1:C4"><filterColumn colId="1"><filters><filter val="East"/></filters></filterColumn></autoFilter>' +
	'<tableColumns count="3">' +
	'<tableColumn id="1" name="Region" totalsRowLabel="Total" headerRowDxfId="1"/>' +
	'<tableColumn id="2" name="Amount_x000a_(USD)" totalsRowFunction="sum" dataDxfId="2" totalsRowDxfId="3"/>' +
	'<tableColumn id="3" name="Tax" totalsRowFunction="custom"><calculatedColumnFormula>Sales_Table[[#This Row],[Amount' +
	'_x000a_(USD)]]*0.2</calculatedColumnFormula><totalsRowFormula>_xlfn.STDEV.S(Sales_Table[Tax])</totalsRowFormula></tableColumn>' +
	'</tableColumns><tableStyleInfo name="TableStyleMedium2" showFirstColumn="1" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/></table>';

const COMMENTS_XML =
	`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<comments xmlns="${MAIN_NS}"><authors><author>Alice</author>` +
	'<author>tc={4C3D0B6B-0000-0000-0000-000000000001}</author></authors><commentList>' +
	'<comment ref="B2" authorId="0" shapeId="0"><text><r><rPr><b/><sz val="9"/><color indexed="81"/><rFont val="Tahoma"/></rPr><t>Alice:</t></r>' +
	'<r><rPr><sz val="9"/><rFont val="Tahoma"/></rPr><t xml:space="preserve">\nCheck this</t></r></text></comment>' +
	'<comment ref="C3" authorId="0"><text><t>Plain note</t></text></comment>' +
	'<comment ref="D4" authorId="1"><text><t>[Threaded comment] mirrored text</t></text></comment>' +
	'</commentList></comments>';

const VML =
	'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
	'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
	'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe"><v:stroke joinstyle="miter"/></v:shapetype>' +
	'<v:shape id="_x0000_s1025" type="#_x0000_t202" style="position:absolute;margin-left:107.25pt;width:144pt;height:79.5pt;z-index:1;visibility:visible" fillcolor="#ffffe1">' +
	'<v:fill color2="#ffffe1"/><v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
	'<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/><x:Anchor>2, 15, 0, 10, 4, 15, 4, 4</x:Anchor><x:AutoFill>False</x:AutoFill>' +
	'<x:Row>1</x:Row><x:Column>1</x:Column><x:Visible/></x:ClientData></v:shape>' +
	"<v:shape id='_x0000_s1026' type='#_x0000_t202' style='position:absolute;width:96px;height:1in;visibility:hidden'>" +
	'<x:ClientData ObjectType="Note"><x:Row>2</x:Row><x:Column>2</x:Column></x:ClientData></v:shape>' +
	'<v:shape id="_x0000_s1027" style="width:10pt"><x:ClientData ObjectType="Button"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>' +
	'</xml>';

const THREADED_XML =
	'<ThreadedComments xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="' +
	MAIN_NS +
	'"><threadedComment ref="D4" dT="2023-05-01T10:15:30.00" personId="{P1}" id="{4C3D0B6B-0000-0000-0000-000000000001}"><text>Looks wrong?</text></threadedComment>' +
	'<threadedComment ref="D4" dT="2023-05-01T11:00:00.00" personId="{P2}" id="{T2}" parentId="{4C3D0B6B-0000-0000-0000-000000000001}" done="1"><text>Fixed &amp; done</text></threadedComment>' +
	'</ThreadedComments>';

const PERSONS_XML =
	'<personList xmlns="http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments" xmlns:x="' +
	MAIN_NS +
	'"><person displayName="Bob Builder" id="{P1}" userId="bob@example.com" providerId="AD"/><person displayName="Carol" id="{P2}" providerId="None"/></personList>';

describe('tables', () => {
	it('reads every table field, its filter, style and round-trip data', async () => {
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'T',
						xml: worksheet('', {
							after: '<tableParts count="1"><tablePart r:id="rIdT"/></tableParts>',
						}),
						rels: [{ id: 'rIdT', type: 'table', target: '../tables/table3.xml' }],
					},
				],
				parts: { 'xl/tables/table3.xml': TABLE_XML },
			}),
		);
		const [table] = wb.sheets[0].tables;
		expect(table).toMatchObject({
			id: 3,
			name: 'Sales',
			displayName: 'Sales_Table',
			ref: { top: 0, left: 0, bottom: 4, right: 2 },
			headerRowCount: 1,
			totalsRowCount: 1,
			totalsRowShown: true,
			headerRowDxfId: 4,
			dataDxfId: 5,
			totalsRowDxfId: 6,
			comment: 'Quarterly',
			partPath: 'xl/tables/table3.xml',
			dirty: false,
			autoFilter: {
				ref: { top: 0, left: 0, bottom: 3, right: 2 },
				columns: [{ colId: 1, filters: { values: ['East'] } }],
			},
			style: {
				name: 'TableStyleMedium2',
				showFirstColumn: true,
				showLastColumn: false,
				showRowStripes: true,
				showColumnStripes: false,
			},
		});
		expect(table.columns).toEqual([
			{ id: 1, name: 'Region', totalsRowLabel: 'Total', headerRowDxfId: 1 },
			{ id: 2, name: 'Amount\n(USD)', totalsRowFunction: 'sum', dataDxfId: 2, totalsRowDxfId: 3 },
			{
				id: 3,
				name: 'Tax',
				totalsRowFunction: 'custom',
				calculatedColumnFormula: 'Sales_Table[[#This Row],[Amount_x000a_(USD)]]*0.2',
				totalsRowFormula: 'STDEV.S(Sales_Table[Tax])',
			},
		]);
		expect(table.raw?.name).toBe('table');
	});
});

describe('comments', () => {
	async function load() {
		return readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'C',
						xml: worksheet('', { after: '<legacyDrawing r:id="rIdV"/>' }),
						rels: [
							{ id: 'rIdC', type: 'comments', target: '../comments1.xml' },
							{ id: 'rIdV', type: 'vmlDrawing', target: '../drawings/vmlDrawing1.vml' },
							{
								id: 'rIdTC',
								type: 'http://schemas.microsoft.com/office/2017/10/relationships/threadedComment',
								target: '../threadedComments/threadedComment1.xml',
							},
						],
					},
				],
				workbookRels: [
					{
						id: 'rIdP',
						type: 'http://schemas.microsoft.com/office/2017/10/relationships/person',
						target: 'persons/person.xml',
					},
				],
				parts: {
					'xl/comments1.xml': COMMENTS_XML,
					'xl/drawings/vmlDrawing1.vml': VML,
					'xl/threadedComments/threadedComment1.xml': THREADED_XML,
					'xl/persons/person.xml': PERSONS_XML,
				},
			}),
		);
	}

	it('reads legacy notes with rich text, VML visibility and size, skipping threaded mirrors', async () => {
		const wb = await load();
		const sheet = wb.sheets[0];
		expect(sheet.notes).toEqual([
			{
				ref: { row: 1, col: 1 },
				author: 'Alice',
				text: 'Alice:\nCheck this',
				runs: [
					{ text: 'Alice:', font: { bold: true, size: 9, color: { indexed: 81 }, name: 'Tahoma' } },
					{ text: '\nCheck this', font: { size: 9, name: 'Tahoma' } },
				],
				visible: true,
				width: 192,
				height: 106,
			},
			{ ref: { row: 2, col: 2 }, author: 'Alice', text: 'Plain note', width: 96, height: 96 },
		]);
		expect(sheet.source).toMatchObject({
			commentsPath: 'xl/comments1.xml',
			vmlPath: 'xl/drawings/vmlDrawing1.vml',
			threadedCommentsPath: 'xl/threadedComments/threadedComment1.xml',
		});
	});

	it('reads threaded comments and the workbook person list', async () => {
		const wb = await load();
		expect(wb.sheets[0].threadedComments).toEqual([
			{
				id: '{4C3D0B6B-0000-0000-0000-000000000001}',
				ref: { row: 3, col: 3 },
				personId: '{P1}',
				text: 'Looks wrong?',
				created: '2023-05-01T10:15:30.00',
			},
			{
				id: '{T2}',
				ref: { row: 3, col: 3 },
				personId: '{P2}',
				text: 'Fixed & done',
				created: '2023-05-01T11:00:00.00',
				parentId: '{4C3D0B6B-0000-0000-0000-000000000001}',
				done: true,
			},
		]);
		expect(wb.persons).toEqual([
			{ id: '{P1}', displayName: 'Bob Builder', userId: 'bob@example.com', providerId: 'AD' },
			{ id: '{P2}', displayName: 'Carol', providerId: 'None' },
		]);
		expect(wb.warnings).toEqual([]);
	});

	it('tolerates VML that is not well-formed XML', () => {
		const shapes = parseVmlNoteShapes(
			'<xml><v:shape style="width:72pt;height:36pt"><div><br><br></div><x:ClientData ObjectType="Note">' +
				'<x:Row>4</x:Row><x:Column>0</x:Column><x:Visible>False</x:Visible></x:ClientData></v:shape>' +
				'<v:shape id="s2"/><v:shape style="width:3cm"><x:ClientData ObjectType="Note"><x:Row>5</x:Row><x:Column>1</x:Column></x:ClientData></v:shape></xml>',
		);
		expect(shapes).toEqual([
			{ row: 4, col: 0, visible: false, width: 96, height: 48 },
			{ row: 5, col: 1, visible: false, width: 113.39 },
		]);
	});

	it('records a warning when the comments part is broken but keeps the sheet', async () => {
		const wb = await readWorkbook(
			await buildWorkbook({
				sheets: [
					{
						name: 'C',
						xml: worksheet('<row r="1"><c r="A1"><v>3</v></c></row>'),
						rels: [{ id: 'rIdC', type: `${REL_BASE}/comments`, target: '../comments1.xml' }],
					},
				],
				parts: { 'xl/comments1.xml': '<comments><commentList>' },
			}),
		);
		expect(wb.sheets[0].cells.size).toBe(1);
		expect(wb.sheets[0].notes).toEqual([]);
		expect(wb.warnings.map((w) => [w.code, w.part])).toEqual([
			['unreadable-part', 'xl/comments1.xml'],
		]);
	});
});
