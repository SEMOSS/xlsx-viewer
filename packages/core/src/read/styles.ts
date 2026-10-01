import { createDefaultStyleSheet } from '../styles/defaults';
import type {
	CustomTableStyle,
	NamedCellStyle,
	StyleSheet,
	StyleSheetRaw,
	TableStyleElement,
} from '../types/style';
import type { XmlElement } from '../xml/types';
import { parseSpreadsheetXml } from './xml-normalize';
import { normalizeArgb, parseBorder, parseDxf, parseFill, parseFont, parseXf } from './style-parts';
import { bool, compact, int, kid, kids, str } from './xml-utils';

/**
 * `xl/styles.xml` -> {@link StyleSheet}. Every collection entry keeps its
 * original element in `styles.raw` so the writer can re-emit unchanged
 * entries verbatim; only entries appended later are serialized from the model.
 */
export function parseStyleSheet(xml: string): StyleSheet {
	return parseStyleSheetElement(parseSpreadsheetXml(xml).root);
}

export function parseStyleSheetElement(root: XmlElement): StyleSheet {
	const defaults = createDefaultStyleSheet();
	const raw: StyleSheetRaw = { root };

	const numFmts: Record<number, string> = {};
	for (const nf of kids(kid(root, 'numFmts'), 'numFmt')) {
		const id = int(nf, 'numFmtId');
		const code = str(nf, 'formatCode');
		if (id !== undefined && code !== undefined) {
			numFmts[id] = code;
		}
	}

	const fontEls = kids(kid(root, 'fonts'), 'font');
	const fillEls = kids(kid(root, 'fills'), 'fill');
	const borderEls = kids(kid(root, 'borders'), 'border');
	const styleXfEls = kids(kid(root, 'cellStyleXfs'), 'xf');
	const cellXfEls = kids(kid(root, 'cellXfs'), 'xf');
	const cellStyleEls = kids(kid(root, 'cellStyles'), 'cellStyle');
	const dxfEls = kids(kid(root, 'dxfs'), 'dxf');
	raw.fonts = fontEls;
	raw.fills = fillEls;
	raw.borders = borderEls;
	raw.cellStyleXfs = styleXfEls;
	raw.cellXfs = cellXfEls;
	raw.cellStyles = cellStyleEls;
	raw.dxfs = dxfEls;

	const styles: StyleSheet = {
		numFmts,
		fonts: fontEls.map((f) => parseFont(f)),
		fills: fillEls.map((f) => parseFill(f)),
		borders: borderEls.map((b) => parseBorder(b)),
		cellStyleXfs: styleXfEls.map((x) => parseXf(x, false)),
		cellXfs: cellXfEls.map((x) => parseXf(x, true)),
		cellStyles: cellStyleEls.map(parseCellStyle),
		dxfs: dxfEls.map((d) => parseDxf(d)),
		raw,
	};

	// A styles part without the mandatory first entries would leave index 0
	// dangling; pad with the defaults (appended entries have no raw XML).
	if (styles.fonts.length === 0) {
		styles.fonts = defaults.fonts;
	}
	if (styles.fills.length === 0) {
		styles.fills = defaults.fills;
	}
	if (styles.borders.length === 0) {
		styles.borders = defaults.borders;
	}
	if (styles.cellStyleXfs.length === 0) {
		styles.cellStyleXfs = defaults.cellStyleXfs;
	}
	if (styles.cellXfs.length === 0) {
		styles.cellXfs = defaults.cellXfs;
	}

	const indexed = kids(kid(kid(root, 'colors'), 'indexedColors'), 'rgbColor');
	if (indexed.length > 0) {
		styles.indexedColors = indexed.map((c) => normalizeArgb(str(c, 'rgb') ?? 'FF000000'));
	}

	const tableStyles = kid(root, 'tableStyles');
	if (tableStyles) {
		const defTable = str(tableStyles, 'defaultTableStyle');
		if (defTable !== undefined) {
			styles.defaultTableStyle = defTable;
		}
		const defPivot = str(tableStyles, 'defaultPivotStyle');
		if (defPivot !== undefined) {
			styles.defaultPivotStyle = defPivot;
		}
		const custom = kids(tableStyles, 'tableStyle').map(parseTableStyle);
		if (custom.length > 0) {
			styles.customTableStyles = custom;
		}
	}
	return styles;
}

function parseCellStyle(el: XmlElement): NamedCellStyle {
	return compact<NamedCellStyle>({
		name: str(el, 'name') ?? '',
		xfId: int(el, 'xfId') ?? 0,
		builtinId: int(el, 'builtinId'),
		iLevel: int(el, 'iLevel'),
		hidden: bool(el, 'hidden'),
		customBuiltin: bool(el, 'customBuiltin'),
	});
}

function parseTableStyle(el: XmlElement): CustomTableStyle {
	const elements: TableStyleElement[] = kids(el, 'tableStyleElement').map((e) =>
		compact<TableStyleElement>({
			type: str(e, 'type') ?? '',
			dxfId: int(e, 'dxfId') ?? 0,
			size: int(e, 'size'),
		}),
	);
	return {
		name: str(el, 'name') ?? '',
		// Both default to true in the schema.
		pivot: bool(el, 'pivot') ?? true,
		table: bool(el, 'table') ?? true,
		elements,
	};
}
