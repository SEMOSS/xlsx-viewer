import { describe, expect, it } from 'vitest';

import { readWorkbook } from '../read';
import { createWorkbook } from '../sheet/workbook';
import type { DrawingObject } from '../types/drawing';
import type { XlsxWarning } from '../types/warnings';
import { parseXml } from '../xml';
import { anchorXml } from './drawing-xml';
import { writeWorkbook } from './index';
import { isSelfContained, prefixScope } from './raw-markup';
import {
	childNames,
	openPackage,
	partRoot,
	pick,
	pickAll,
	validatePackage,
} from './test-utils/package-check';

const PNG = new Uint8Array(
	Buffer.from(
		'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNkYPhfDwAChwGA60e6kgAAAABJRU5ErkJggg==',
		'base64',
	),
);
const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
const at = (col: number, row: number) => ({
	type: 'twoCell' as const,
	from: { col, colOff: 0, row, rowOff: 0 },
	to: { col: col + 2, colOff: 0, row: row + 2, rowOff: 0 },
});

async function drawingOf(objects: DrawingObject[], media: Record<string, Uint8Array> = {}) {
	const wb = createWorkbook();
	for (const [path, data] of Object.entries(media)) {
		wb.media.set(path, { path, contentType: 'image/png', data });
	}
	wb.sheets[0].drawings = objects;
	const warnings: XlsxWarning[] = [];
	const pkg = await openPackage(await writeWorkbook(wb, { onWarning: (w) => warnings.push(w) }));
	expect(await validatePackage(pkg)).toEqual([]);
	const root = pkg.has('xl/drawings/drawing1.xml')
		? await partRoot(pkg, 'xl/drawings/drawing1.xml')
		: undefined;
	return { pkg, root, warnings, rels: await pkg.rels('xl/drawings/drawing1.xml') };
}

describe('anchors', () => {
	it('serializes the three anchor kinds', () => {
		const wrap = (xml: string) => parseXml(`<r xmlns:xdr="${XDR}">${xml}</r>`).root;
		const two = wrap(anchorXml({ ...at(1, 2), editAs: 'oneCell' }, ''));
		expect(pick(two, 'twoCellAnchor')?.attrs.editAs).toBe('oneCell');
		expect(childNames(pick(two, 'twoCellAnchor'))).toEqual(['from', 'to', 'clientData']);
		const one = wrap(
			anchorXml(
				{
					type: 'oneCell',
					from: { col: 0, colOff: 5.4, row: -1, rowOff: 0 },
					ext: { cx: 10, cy: 20 },
				},
				'',
			),
		);
		expect(pick(one, 'oneCellAnchor', 'from', 'colOff')?.children[0]).toBe('5');
		expect(pick(one, 'oneCellAnchor', 'from', 'row')?.children[0]).toBe('0');
		expect(pick(one, 'oneCellAnchor', 'ext')?.attrs).toEqual({ cx: '10', cy: '20' });
		const abs = wrap(
			anchorXml({ type: 'absolute', pos: { x: 1, y: 2 }, ext: { cx: 3, cy: 4 } }, ''),
		);
		expect(childNames(pick(abs, 'absoluteAnchor'))).toEqual(['pos', 'ext', 'clientData']);
	});
});

describe('drawing parts', () => {
	it('writes cropped, hyperlinked pictures and stores media under a package path', async () => {
		const { pkg, root, rels } = await drawingOf(
			[
				{
					kind: 'picture',
					id: '1',
					mediaPath: 'blob:abc',
					anchor: at(0, 0),
					crop: { l: 1000, r: 2000 },
					hyperlink: 'https://example.com',
					description: 'Alt',
				},
				{
					kind: 'picture',
					id: '1',
					mediaPath: 'blob:abc',
					anchor: at(4, 0),
					hyperlink: '#Sheet1!A1',
				},
			],
			{ 'blob:abc': PNG },
		);
		expect(pkg.partPaths().filter((p) => p.startsWith('xl/media/'))).toEqual([
			'xl/media/image1.png',
		]);
		const pics = pickAll(root, 'twoCellAnchor').map((a) => pick(a, 'pic')!);
		const ids = pics.map((p) => pick(p, 'nvPicPr', 'cNvPr')!.attrs.id);
		expect(new Set(ids).size).toBe(2);
		expect(pick(pics[0], 'nvPicPr', 'cNvPr')?.attrs.descr).toBe('Alt');
		expect(pick(pics[0], 'blipFill', 'srcRect')?.attrs).toEqual({ l: '1000', r: '2000' });
		const embeds = pics.map((p) => pick(p, 'blipFill', 'blip')!.attrs['r:embed']);
		expect(embeds[0]).toBe(embeds[1]);
		expect(rels.find((r) => r.id === embeds[0])?.target).toBe('../media/image1.png');
		const links = rels.filter((r) => r.type.endsWith('/hyperlink'));
		expect(links.map((r) => [r.target, r.targetMode])).toEqual([
			['https://example.com', 'External'],
			['#Sheet1!A1', undefined],
		]);
	});

	it('skips pictures without image data and drops an empty drawing', async () => {
		const { pkg, warnings } = await drawingOf([
			{ kind: 'picture', id: '1', mediaPath: 'xl/media/missing.png', anchor: at(0, 0) },
		]);
		expect(pkg.has('xl/drawings/drawing1.xml')).toBe(false);
		expect(pick(await partRoot(pkg, 'xl/worksheets/sheet1.xml'), 'drawing')).toBeUndefined();
		expect(warnings.map((w) => w.code)).toEqual(['missing-media']);
	});

	it('writes shapes, connectors and groups', async () => {
		const { root } = await drawingOf([
			{
				kind: 'group',
				id: '10',
				anchor: at(1, 1),
				childOffset: { x: 0, y: 0 },
				childExtent: { cx: 1000, cy: 1000 },
				children: [
					{
						kind: 'shape',
						id: '11',
						geometry: 'rect',
						anchor: { type: 'absolute', pos: { x: 0, y: 0 }, ext: { cx: 500, cy: 500 } },
						rotation: 45,
						flipH: true,
					},
					{
						kind: 'shape',
						id: '12',
						geometry: 'line',
						isConnector: true,
						anchor: { type: 'absolute', pos: { x: 500, y: 500 }, ext: { cx: 500, cy: 500 } },
						line: { width: 9525, tailEnd: { type: 'triangle' } },
					},
				],
			},
			{
				kind: 'shape',
				id: '13',
				geometry: 'custom',
				adjust: { adj: 5000 },
				anchor: at(5, 5),
				fill: {
					type: 'gradient',
					angle: 45,
					stops: [
						{
							pos: 0,
							color: {
								kind: 'scheme',
								value: 'accent1',
								transforms: [{ name: 'lumMod', val: 75000 }],
							},
						},
						{ pos: 1, color: { kind: 'srgb', value: 'FFFFFF' } },
					],
				},
			},
		]);
		const [groupAnchor, shapeAnchor] = pickAll(root, 'twoCellAnchor');
		const group = pick(groupAnchor, 'grpSp')!;
		expect(childNames(group)).toEqual(['nvGrpSpPr', 'grpSpPr', 'sp', 'cxnSp']);
		expect(childNames(pick(group, 'grpSpPr', 'xfrm'))).toEqual(['off', 'ext', 'chOff', 'chExt']);
		expect(pick(group, 'grpSpPr', 'xfrm', 'chExt')?.attrs).toEqual({ cx: '1000', cy: '1000' });
		const child = pick(group, 'sp', 'spPr', 'xfrm')!;
		expect(child.attrs).toEqual({ rot: '2700000', flipH: '1' });
		expect(pick(child, 'ext')?.attrs).toEqual({ cx: '500', cy: '500' });
		const cxn = pick(group, 'cxnSp')!;
		expect(childNames(cxn)).toEqual(['nvCxnSpPr', 'spPr', 'style']);
		expect(pick(cxn, 'spPr', 'ln', 'tailEnd')?.attrs.type).toBe('triangle');
		const shape = pick(shapeAnchor, 'sp')!;
		expect(pick(shape, 'spPr', 'prstGeom')?.attrs.prst).toBe('rect');
		const grad = pick(shape, 'spPr', 'gradFill')!;
		expect(pickAll(pick(grad, 'gsLst'), 'gs').map((g) => g.attrs.pos)).toEqual(['0', '100000']);
		expect(pick(grad, 'lin')?.attrs.ang).toBe('2700000');
		expect(pick(grad, 'gsLst', 'gs', 'schemeClr', 'lumMod')?.attrs.val).toBe('75000');
		expect(childNames(shape)).toEqual(['nvSpPr', 'spPr', 'style']);
	});

	it('re-anchors edited unknown objects and keeps alternate content verbatim', async () => {
		const raw = parseXml(
			`<xdr:twoCellAnchor xmlns:xdr="${XDR}"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>1</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="7" name="Ole"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>`,
		).root;
		const alt = parseXml(
			`<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"><mc:Choice xmlns:a14="urn:a14" Requires="a14"><xdr:twoCellAnchor xmlns:xdr="${XDR}"><xdr:from><xdr:col>3</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>4</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>4</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:sp><xdr:nvSpPr><xdr:cNvPr id="8" name="Ink"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr/></xdr:sp><xdr:clientData/></xdr:twoCellAnchor></mc:Choice></mc:AlternateContent>`,
		).root;
		const { root } = await drawingOf([
			{ kind: 'unknown', id: '7', anchor: at(6, 6), raw, dirty: true },
			{ kind: 'unknown', id: '8', anchor: at(3, 3), raw: alt, dirty: true },
		]);
		const moved = pick(root, 'twoCellAnchor')!;
		expect(pick(moved, 'from', 'col')?.children[0]).toBe('6');
		expect(pick(moved, 'sp', 'nvSpPr', 'cNvPr')?.attrs.name).toBe('Ole');
		expect(
			pick(root, 'AlternateContent', 'Choice', 'twoCellAnchor', 'sp', 'nvSpPr', 'cNvPr')?.attrs
				.name,
		).toBe('Ink');
	});
});

describe('reusing original markup', () => {
	it('accepts only subtrees whose prefixes resolve and that reference no relationships', () => {
		const scope = prefixScope({ 'xmlns:xdr': XDR, 'xmlns:a': 'urn:a' });
		const ok = parseXml(`<xdr:sp xmlns:a14="urn:a14"><a:t a14:x="1"/></xdr:sp>`).root;
		const undeclared = parseXml('<xdr:sp><mc:Fallback/></xdr:sp>').root;
		const linked = parseXml(
			'<xdr:pic xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><a:blip r:embed="rId1"/></xdr:pic>',
		).root;
		expect(isSelfContained(ok, scope)).toBe(true);
		expect(isSelfContained(undeclared, scope)).toBe(false);
		expect(isSelfContained(linked, scope)).toBe(false);
	});

	it('regenerates loaded pictures when their sheet is copied into fresh parts', async () => {
		const wb = createWorkbook();
		wb.media.set('xl/media/image1.png', {
			path: 'xl/media/image1.png',
			contentType: 'image/png',
			data: PNG,
		});
		wb.sheets[0].drawings = [
			{ kind: 'picture', id: '2', mediaPath: 'xl/media/image1.png', anchor: at(1, 1) },
		];
		const loaded = await readWorkbook(await writeWorkbook(wb));
		const original = loaded.sheets[0];
		expect(original.drawings[0].dirty).toBe(false);
		loaded.sheets.push({
			...original,
			id: 'copy',
			name: 'Copy',
			sheetId: 9,
			cells: new Map(original.cells),
		});
		loaded.dirty = true;
		const pkg = await openPackage(await writeWorkbook(loaded));
		expect(await validatePackage(pkg)).toEqual([]);
		const drawings = pkg.partPaths().filter((p) => /^xl\/drawings\/drawing\d+\.xml$/.test(p));
		expect(drawings).toHaveLength(2);
		for (const d of drawings) {
			const blip = pick(await partRoot(pkg, d), 'twoCellAnchor', 'pic', 'blipFill', 'blip')!;
			const rel = (await pkg.rels(d)).find((r) => r.id === blip.attrs['r:embed']);
			expect(rel?.target).toBe('../media/image1.png');
		}
		const back = await readWorkbook(await writeWorkbook(loaded));
		expect(back.sheets.map((s) => s.drawings.length)).toEqual([1, 1]);
	});
});
