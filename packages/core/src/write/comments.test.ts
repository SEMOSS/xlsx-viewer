import { describe, expect, it } from 'vitest';

import { createWorksheet } from '../sheet/worksheet';
import { parseXml } from '../xml';
import { buildCommentsXml, legacyComments, mirrorText, threadsOf } from './comments';
import { SheetGeometry } from './sheet-geometry';
import { childNames, pick, pickAll } from './test-utils/package-check';
import { buildPersonsXml, buildThreadedCommentsXml, GuidMap } from './threaded-comments';
import { buildVmlXml, noteAnchor, parseOriginalVml, VmlBlockAllocator, vmlBlocks } from './vml';

const T1 = '{00000000-0000-0000-0000-00000000000A}';

describe('legacy comments', () => {
	it('merges notes with thread mirrors (one comment per cell)', () => {
		const s = createWorksheet('S', 1);
		s.notes = [
			{ ref: { row: 3, col: 0 }, author: 'Ann', text: 'later' },
			{
				ref: { row: 0, col: 1 },
				author: 'Ann',
				text: 'rich',
				runs: [{ text: 'ri', font: { bold: true } }, { text: 'ch' }],
			},
			{ ref: { row: 0, col: 2 }, author: 'Cy', text: 'hidden by thread' },
		];
		s.threadedComments = [
			{ id: T1, ref: { row: 0, col: 2 }, personId: 'p', text: 'Q' },
			{ id: 'r', ref: { row: 0, col: 2 }, personId: 'p', text: 'A', parentId: T1 },
		];
		const { comments, droppedNotes } = legacyComments(s, new Map());
		expect(droppedNotes).toBe(1);
		expect(comments.map((c) => [c.ref.row, c.ref.col, c.author])).toEqual([
			[0, 1, 'Ann'],
			[0, 2, `tc=${T1}`],
			[3, 0, 'Ann'],
		]);
		expect(comments[1].text).toBe(mirrorText(s.threadedComments[0], [s.threadedComments[1]]));
		expect(comments[1].text.startsWith('[Threaded comment]\n\nYour version of Excel')).toBe(true);
		expect(threadsOf(s.threadedComments)[0].replies).toHaveLength(1);

		const root = parseXml(buildCommentsXml(comments)).root;
		expect(root.attrs['mc:Ignorable']).toBe('xr');
		expect(pickAll(pick(root, 'authors'), 'author').map((a) => a.children[0])).toEqual([
			'Ann',
			`tc=${T1}`,
		]);
		const list = pickAll(pick(root, 'commentList'), 'comment');
		expect(list.map((c) => c.attrs)).toEqual([
			{ ref: 'B1', authorId: '0', shapeId: '0' },
			{ ref: 'C1', authorId: '1', shapeId: '0', 'xr:uid': T1 },
			{ ref: 'A4', authorId: '0', shapeId: '0' },
		]);
		expect(childNames(pick(list[0], 'text'))).toEqual(['r', 'r']);
		expect(pick(list[2], 'text', 't')?.children[0]).toBe('later');
	});

	it('omits the revision namespace when no comment has a uid', () => {
		const xml = buildCommentsXml([{ ref: { row: 0, col: 0 }, author: 'A & B', text: 'x' }]);
		expect(xml).not.toContain('xmlns:xr');
		expect(xml).toContain('<author>A &amp; B</author>');
	});
});

describe('VML notes', () => {
	const sheet = createWorksheet('S', 1);
	const g = new SheetGeometry(sheet);

	it('uses Excel default anchors and sizes boxes from known pixel sizes', () => {
		expect(noteAnchor(g, { row: 0, col: 0 })).toEqual([1, 15, 0, 2, 3, 15, 4, 16]);
		expect(noteAnchor(g, { row: 2, col: 1 })).toEqual([2, 15, 1, 10, 4, 15, 5, 16]);
		// 64px columns, 20px rows: 15 + 144 = 159px -> col 4 offset 31; 10 + 79 = 89px -> row 5 offset 9.
		expect(noteAnchor(g, { row: 2, col: 1, width: 144, height: 79 })).toEqual([
			2, 15, 1, 10, 4, 31, 5, 9,
		]);
	});

	it('writes note shapes in fresh idmap blocks', () => {
		const xml = buildVmlXml(
			[
				{ row: 0, col: 0 },
				{ row: 4, col: 2, visible: true },
			],
			g,
			[3],
		);
		const root = parseXml(xml).root;
		expect(root.name).toBe('xml');
		expect(xml.startsWith('<xml')).toBe(true);
		expect(pick(root, 'shapelayout', 'idmap')?.attrs.data).toBe('3');
		const shapes = pickAll(root, 'shape');
		expect(shapes.map((s) => s.attrs.id)).toEqual(['_x0000_s3073', '_x0000_s3074']);
		expect(shapes[1].attrs.style).toContain('visibility:visible');
		expect(childNames(pick(shapes[1], 'ClientData'))).toContain('Visible');
		expect(pick(shapes[1], 'ClientData', 'Row')?.children[0]).toBe('4');
		expect(pickAll(root, 'shapetype').map((t) => t.attrs.id)).toEqual(['_x0000_t202']);
	});

	it('keeps form controls of the original drawing', () => {
		const original =
			'<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">' +
			'<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>' +
			'<v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"/>' +
			'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"/>' +
			'<v:shape id="_x0000_s1025" type="#_x0000_t201"><x:ClientData ObjectType="Checkbox"><x:Anchor>1, 1, 1, 1, 2, 1, 2, 1</x:Anchor></x:ClientData></v:shape>' +
			'<v:shape id="_x0000_s1026" type="#_x0000_t202"><x:ClientData ObjectType="Note"><x:Row>0</x:Row><x:Column>0</x:Column></x:ClientData></v:shape>' +
			'</xml>';
		expect(vmlBlocks(original)).toEqual([1]);
		const parsed = parseOriginalVml(original)!;
		expect(parsed.hasNoteType).toBe(true);
		const root = parseXml(buildVmlXml([{ row: 5, col: 5 }], g, [2], parsed)).root;
		expect(pick(root, 'shapelayout', 'idmap')?.attrs.data).toBe('1,2');
		const shapes = pickAll(root, 'shape');
		expect(shapes.map((s) => s.attrs.id)).toEqual(['_x0000_s1025', '_x0000_s2049']);
		expect(pickAll(root, 'shapetype')).toHaveLength(2);
		expect(parseOriginalVml('<xml><br></xml>')).toBeUndefined();
	});

	it('allocates unique blocks across the workbook', () => {
		const alloc = new VmlBlockAllocator();
		alloc.reserve([1, 2]);
		expect(alloc.allocate(5)).toEqual([3]);
		expect(alloc.allocate(2000)).toEqual([4, 5]);
	});
});

describe('threaded comments', () => {
	it('writes GUID ids consistently and the person list', () => {
		const ids = new GuidMap();
		const persons = new GuidMap();
		const now = new Date('2024-03-04T05:06:07.891Z');
		const xml = buildThreadedCommentsXml(
			[
				{ id: T1, ref: { row: 0, col: 0 }, personId: 'local-user', text: 'Hi & bye', done: true },
				{
					id: 'x',
					ref: { row: 0, col: 0 },
					personId: 'local-user',
					text: 'Re',
					parentId: T1,
					created: '2024-01-01T10:00:00.00',
				},
			],
			ids,
			persons,
			now,
		);
		const items = pickAll(parseXml(xml).root, 'threadedComment');
		expect(items[0].attrs).toMatchObject({
			ref: 'A1',
			id: T1,
			done: '1',
			dT: '2024-03-04T05:06:07.89',
		});
		expect(items[1].attrs.parentId).toBe(T1);
		expect(items[1].attrs.id).toBe(ids.get('x'));
		expect(items[1].attrs.dT).toBe('2024-01-01T10:00:00.00');
		expect(pick(items[0], 'text')?.children[0]).toBe('Hi & bye');
		const personXml = parseXml(
			buildPersonsXml([{ id: 'local-user', displayName: 'Me' }], persons),
		).root;
		expect(pick(personXml, 'person')?.attrs).toEqual({
			displayName: 'Me',
			id: items[0].attrs.personId,
			providerId: 'None',
		});
	});
});
