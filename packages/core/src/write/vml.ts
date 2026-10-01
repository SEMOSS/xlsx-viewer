/**
 * Legacy VML drawing for notes (`xl/drawings/vmlDrawingN.vml`): the standard
 * Excel note shapes (`_x0000_t202`) with anchors computed from the cell.
 * Non-note shapes of the original drawing (form controls) are kept.
 */
import { MAX_COL_INDEX, MAX_ROW_INDEX } from '../address/constants';
import { elements, isElement, localName, parseXml, serializeElement } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlNode } from '../xml/types';
import type { SheetGeometry } from './sheet-geometry';
import { attrString } from './xml-utils';

export interface VmlNote {
	row: number;
	col: number;
	visible?: boolean;
	/** Box size in pixels, when known. */
	width?: number;
	height?: number;
}

/** What survives from the original VML part. */
export interface VmlOriginal {
	rootAttrs: Record<string, string>;
	/** Shape types and non-note shapes, verbatim. */
	kept: XmlNode[];
	/** `o:idmap` blocks of the original part (the kept shapes' ids live there). */
	blocks: number[];
	/** The original defines the note shape type `_x0000_t202`. */
	hasNoteType: boolean;
}

/** Shapes per idmap block (ids `block * 1024 + 1 .. + 1023`). */
export const SHAPES_PER_BLOCK = 1023;

const NOTE_SHAPE_TYPE =
	'<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">' +
	'<v:stroke joinstyle="miter"/><v:path gradientshapeok="t" o:connecttype="rect"/></v:shapetype>';

/** `o:idmap` block numbers declared in a VML part. */
export function vmlBlocks(xml: string): number[] {
	const out: number[] = [];
	const re = /<o:idmap\b[^>]*\bdata="([^"]*)"/g;
	let m: RegExpExecArray | null;
	while ((m = re.exec(xml)) !== null) {
		for (const s of m[1].split(',')) {
			const n = Number(s.trim());
			if (Number.isInteger(n) && n > 0) {
				out.push(n);
			}
		}
	}
	return out;
}

/** Parse the original VML; `undefined` when it is not well-formed XML. */
export function parseOriginalVml(xml: string): VmlOriginal | undefined {
	let root;
	try {
		root = parseXml(xml).root;
	} catch {
		return undefined;
	}
	const kept: XmlNode[] = [];
	let hasNoteType = false;
	for (const c of elements(root)) {
		const name = localName(c.name);
		if (name === 'shapelayout') {
			continue;
		}
		if (name === 'shape') {
			const data = elements(c).find((k) => localName(k.name) === 'ClientData');
			if (data?.attrs.ObjectType === 'Note') {
				continue;
			}
		}
		if (name === 'shapetype' && c.attrs.id === '_x0000_t202') {
			hasNoteType = true;
		}
		kept.push(c);
	}
	return { rootAttrs: { ...root.attrs }, kept, blocks: vmlBlocks(xml), hasNoteType };
}

/** The original has shapes other than notes (controls) that must stay. */
export function hasKeptShapes(original: VmlOriginal | undefined): boolean {
	return !!original?.kept.some((n) => isElement(n) && localName(n.name) !== 'shapetype');
}

/** Workbook-wide allocator of unique idmap blocks. */
export class VmlBlockAllocator {
	private readonly used = new Set<number>();

	reserve(blocks: readonly number[]): void {
		for (const b of blocks) {
			this.used.add(b);
		}
	}

	allocate(shapeCount: number): number[] {
		const need = Math.max(1, Math.ceil(shapeCount / SHAPES_PER_BLOCK));
		const out: number[] = [];
		for (let b = 1; out.length < need; b++) {
			if (!this.used.has(b)) {
				this.used.add(b);
				out.push(b);
			}
		}
		return out;
	}
}

/** Excel's note anchor: `LeftCol, LeftOff, TopRow, TopOff, RightCol, RightOff, BottomRow, BottomOff`. */
export function noteAnchor(g: SheetGeometry, note: VmlNote): number[] {
	const leftCol = Math.min(note.col + 1, MAX_COL_INDEX);
	const topRow = note.row === 0 ? 0 : note.row - 1;
	const topOff = note.row === 0 ? 2 : 10;
	if (note.width === undefined && note.height === undefined) {
		return [
			leftCol,
			15,
			topRow,
			topOff,
			Math.min(leftCol + 2, MAX_COL_INDEX),
			15,
			Math.min(topRow + 4, MAX_ROW_INDEX),
			16,
		];
	}
	const right = g.walkCols(leftCol, 15, note.width ?? 144);
	const bottom = g.walkRows(topRow, topOff, note.height ?? 79);
	return [leftCol, 15, topRow, topOff, right.index, right.offset, bottom.index, bottom.offset];
}

function pt(px: number): string {
	return String(Math.round(px * 0.75 * 100) / 100);
}

function noteShapeXml(note: VmlNote, id: number, z: number, g: SheetGeometry): string {
	const a = noteAnchor(g, note);
	const x = g.colLeft(a[0]) + a[1];
	const y = g.rowTop(a[2]) + a[3];
	const w = note.width ?? Math.max(1, g.colLeft(a[4]) + a[5] - x);
	const h = note.height ?? Math.max(1, g.rowTop(a[6]) + a[7] - y);
	const style =
		`position:absolute;margin-left:${pt(x)}pt;margin-top:${pt(y)}pt;width:${pt(w)}pt;height:${pt(h)}pt;` +
		`z-index:${z};visibility:${note.visible ? 'visible' : 'hidden'}`;
	return (
		`<v:shape id="_x0000_s${id}" type="#_x0000_t202" style="${style}" fillcolor="#ffffe1" o:insetmode="auto">` +
		'<v:fill color2="#ffffe1"/><v:shadow on="t" color="black" obscured="t"/><v:path o:connecttype="none"/>' +
		'<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>' +
		'<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>' +
		`<x:Anchor>${a.join(', ')}</x:Anchor><x:AutoFill>False</x:AutoFill>` +
		`<x:Row>${note.row}</x:Row><x:Column>${note.col}</x:Column>${note.visible ? '<x:Visible/>' : ''}` +
		'</x:ClientData></v:shape>'
	);
}

/** Serialize the VML part (no XML declaration, like Excel). */
export function buildVmlXml(
	notes: readonly VmlNote[],
	g: SheetGeometry,
	blocks: readonly number[],
	original?: VmlOriginal,
): string {
	const rootAttrs: Record<string, string> = {
		'xmlns:v': NS.vml,
		'xmlns:o': NS.office,
		'xmlns:x': NS.excel,
		...(original?.rootAttrs ?? {}),
	};
	const keptShapes = hasKeptShapes(original);
	const idBlocks = [...(keptShapes ? original!.blocks : []), ...blocks];
	const parts: string[] = [
		`<xml${attrString(rootAttrs)}>`,
		`<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="${idBlocks.join(',')}"/></o:shapelayout>`,
	];
	if (notes.length && !original?.hasNoteType) {
		parts.push(NOTE_SHAPE_TYPE);
	}
	for (const node of original?.kept ?? []) {
		parts.push(isElement(node) ? serializeElement(node) : '');
	}
	notes.forEach((note, i) => {
		const block = blocks[Math.floor(i / SHAPES_PER_BLOCK)];
		const id = block * 1024 + (i % SHAPES_PER_BLOCK) + 1;
		parts.push(noteShapeXml(note, id, i + 1, g));
	});
	parts.push('</xml>');
	return parts.join('');
}
