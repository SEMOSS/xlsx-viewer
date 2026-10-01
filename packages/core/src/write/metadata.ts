/**
 * Cell metadata (`xl/metadata.xml`) bookkeeping. Dynamic array formulas
 * carry `cm` pointing at an XLDAPR cell-metadata record; the source record
 * is reused when present, otherwise one is added (or the part is created).
 */
import {
	attrNum,
	el,
	elements,
	localName,
	parseXml,
	serializeElement,
	XML_DECLARATION,
} from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { setOrderedChild } from './xml-utils';

export const CT_SHEET_METADATA =
	'application/vnd.openxmlformats-officedocument.spreadsheetml.sheetMetadata+xml';
export const XDA_NS = 'http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray';
export const DYNAMIC_ARRAY_EXT_URI = '{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}';

const METADATA_ORDER = [
	'metadataTypes',
	'metadataStrings',
	'mdxMetadata',
	'futureMetadata',
	'cellMetadata',
	'valueMetadata',
	'extLst',
] as const;

function byLocal(parent: XmlElement | undefined, local: string): XmlElement[] {
	return elements(parent).filter((c) => localName(c.name) === local);
}

/** The standard XLDAPR metadata type Excel writes. */
function xldaprType(): XmlElement {
	return el('metadataType', {
		name: 'XLDAPR',
		minSupportedVersion: 120000,
		copy: 1,
		pasteAll: 1,
		pasteValues: 1,
		merge: 1,
		splitFirst: 1,
		rowColShift: 1,
		clearFormats: 1,
		clearComments: 1,
		assign: 1,
		coerce: 1,
		cellMeta: 1,
	});
}

function dynamicBk(): XmlElement {
	return el('bk', {}, [
		el('extLst', {}, [
			el('ext', { uri: DYNAMIC_ARRAY_EXT_URI, 'xmlns:xda': XDA_NS }, [
				el('xda:dynamicArrayProperties', { fDynamic: 1, fCollapsed: 0 }),
			]),
		]),
	]);
}

function setCount(parent: XmlElement, local: string): void {
	parent.attrs.count = String(byLocal(parent, local).length);
}

export class CellMetadata {
	private root?: XmlElement;
	private dynamicCm?: number;
	private cellCount = 0;
	private valueCount = 0;
	/** The part content changed (an XLDAPR record was added or the part is new). */
	changed = false;

	/** Package path of the source metadata part, when the workbook had one. */
	constructor(readonly sourcePath?: string) {}

	static parse(xml: string, sourcePath: string): CellMetadata {
		const md = new CellMetadata(sourcePath);
		md.root = parseXml(xml).root;
		md.scan();
		return md;
	}

	private scan(): void {
		const root = this.root!;
		const types = byLocal(byLocal(root, 'metadataTypes')[0], 'metadataType');
		const typeIndex = types.findIndex((t) => t.attrs.name === 'XLDAPR') + 1;
		const cellMeta = byLocal(root, 'cellMetadata')[0];
		const cellBks = byLocal(cellMeta, 'bk');
		this.cellCount = cellBks.length;
		this.valueCount = byLocal(byLocal(root, 'valueMetadata')[0], 'bk').length;
		if (typeIndex === 0) {
			return;
		}
		const future = byLocal(root, 'futureMetadata').find((f) => f.attrs.name === 'XLDAPR');
		const futureBks = byLocal(future, 'bk');
		const dynamicIdx = futureBks.findIndex((bk) => {
			const props = findLocal(bk, 'dynamicArrayProperties');
			return (
				props !== undefined && (props.attrs.fDynamic === '1' || props.attrs.fDynamic === 'true')
			);
		});
		if (dynamicIdx < 0) {
			return;
		}
		const cm = cellBks.findIndex((bk) =>
			byLocal(bk, 'rc').some(
				(rc) => attrNum(rc, 't') === typeIndex && attrNum(rc, 'v') === dynamicIdx,
			),
		);
		if (cm >= 0) {
			this.dynamicCm = cm + 1;
		}
	}

	/** `cm` refers to an existing cell-metadata record. */
	validCm(cm: number): boolean {
		return Number.isInteger(cm) && cm >= 1 && cm <= this.cellCount;
	}

	/** `vm` refers to an existing value-metadata record. */
	validVm(vm: number): boolean {
		return Number.isInteger(vm) && vm >= 1 && vm <= this.valueCount;
	}

	/** `cm` for a dynamic array anchor: the cell's own when valid, else the XLDAPR record. */
	dynamicArrayCm(existing: number | undefined): number {
		if (existing !== undefined && this.validCm(existing)) {
			return existing;
		}
		if (this.dynamicCm === undefined) {
			this.addDynamicRecord();
		}
		return this.dynamicCm!;
	}

	private addDynamicRecord(): void {
		this.changed = true;
		if (!this.root) {
			this.root = el('metadata', { xmlns: NS.main, 'xmlns:xda': XDA_NS });
		}
		const root = this.root;
		let typesEl = byLocal(root, 'metadataTypes')[0];
		if (!typesEl) {
			typesEl = el('metadataTypes');
			setOrderedChild(root, 'metadataTypes', typesEl, METADATA_ORDER);
		}
		let typeIndex =
			byLocal(typesEl, 'metadataType').findIndex((t) => t.attrs.name === 'XLDAPR') + 1;
		if (typeIndex === 0) {
			typesEl.children.push(xldaprType());
			typeIndex = byLocal(typesEl, 'metadataType').length;
		}
		setCount(typesEl, 'metadataType');
		let future = byLocal(root, 'futureMetadata').find((f) => f.attrs.name === 'XLDAPR');
		if (!future) {
			future = el('futureMetadata', { name: 'XLDAPR' });
			const existing = byLocal(root, 'futureMetadata');
			const anchor = existing[existing.length - 1];
			if (anchor) {
				root.children.splice(root.children.indexOf(anchor) + 1, 0, future);
			} else {
				setOrderedChild(root, 'futureMetadata', future, METADATA_ORDER);
			}
		}
		future.children.push(dynamicBk());
		setCount(future, 'bk');
		const v = byLocal(future, 'bk').length - 1;
		let cellMeta = byLocal(root, 'cellMetadata')[0];
		if (!cellMeta) {
			cellMeta = el('cellMetadata');
			setOrderedChild(root, 'cellMetadata', cellMeta, METADATA_ORDER);
		}
		cellMeta.children.push(el('bk', {}, [el('rc', { t: typeIndex, v })]));
		setCount(cellMeta, 'bk');
		this.cellCount = byLocal(cellMeta, 'bk').length;
		this.dynamicCm = this.cellCount;
	}

	/** The part XML (only meaningful once a record exists). */
	serialize(): string {
		return XML_DECLARATION + serializeElement(this.root ?? el('metadata', { xmlns: NS.main }));
	}
}

function findLocal(root: XmlElement, local: string): XmlElement | undefined {
	for (const c of elements(root)) {
		if (localName(c.name) === local) {
			return c;
		}
		const deep = findLocal(c, local);
		if (deep) {
			return deep;
		}
	}
	return undefined;
}
