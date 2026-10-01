/**
 * BIFF formatting records -> `StyleSheet`. BIFF keeps style XFs and cell
 * XFs in one table; SpreadsheetML splits them into `cellStyleXfs` and
 * `cellXfs`, so XF indices are remapped. Fonts keep the file order (BIFF
 * font index 4 does not exist, so indices above it shift down by one);
 * fills and borders are deduplicated.
 */
import { DEFAULT_INDEXED_COLORS } from '../../color/indexed-palette';
import { BUILTIN_NUMBER_FORMATS } from '../../styles/builtin-number-formats';
import type {
	BorderSpec,
	CellXf,
	FillSpec,
	FontSpec,
	NamedCellStyle,
	StyleSheet,
} from '../../types/style';
import type { RawStyle, RawXf } from './xf-records';

type ApplyFlag =
	| 'applyNumberFormat'
	| 'applyFont'
	| 'applyFill'
	| 'applyBorder'
	| 'applyAlignment'
	| 'applyProtection';

export interface XlsStyleInput {
	/** FONT records in file order. */
	fonts: FontSpec[];
	/** FORMAT records: number format code by id. */
	formats: Map<number, string>;
	xfs: RawXf[];
	styles: RawStyle[];
	/** PALETTE colors (`RRGGBB`) for indices 8 and up. */
	palette?: string[];
}

export interface XlsStyles {
	sheet: StyleSheet;
	/** `cellXfs` index for a BIFF XF index (0 for unknown indices). */
	cellXf(ixfe: number): number;
	/** Font for a BIFF font index (as used by XF records and rich-text runs). */
	font(ifnt: number): FontSpec | undefined;
	/** The cell referred to a missing XF. */
	badXf: boolean;
}

/** Built-in cell style names by `istyBuiltIn` (RowLevel / ColLevel get the level appended). */
function builtinStyleName(id: number, level: number): string | undefined {
	const fixed: Record<number, string> = {
		0: 'Normal',
		3: 'Comma',
		4: 'Currency',
		5: 'Percent',
		6: 'Comma [0]',
		7: 'Currency [0]',
		8: 'Hyperlink',
		9: 'Followed Hyperlink',
		10: 'Note',
		11: 'Warning Text',
		15: 'Title',
		16: 'Heading 1',
		17: 'Heading 2',
		18: 'Heading 3',
		19: 'Heading 4',
		20: 'Input',
		21: 'Output',
		22: 'Calculation',
		23: 'Check Cell',
		24: 'Linked Cell',
		25: 'Total',
		26: 'Good',
		27: 'Bad',
		28: 'Neutral',
		53: 'Explanatory Text',
	};
	if (id === 1 || id === 2) {
		return `${id === 1 ? 'RowLevel' : 'ColLevel'}_${level + 1}`;
	}
	if (id >= 29 && id <= 52) {
		const accent = Math.floor((id - 29) / 4) + 1;
		const shade = ['', '20% - ', '40% - ', '60% - '][(id - 29) % 4];
		return `${shade}Accent${accent}`;
	}
	return fixed[id];
}

/** Model index of a BIFF font index (BIFF skips index 4). */
function fontIndex(ifnt: number, count: number): number {
	const i = ifnt >= 4 ? ifnt - 1 : ifnt;
	return ifnt !== 4 && i < count ? i : 0;
}

/** Build the stylesheet and the XF index maps. */
export function buildStyles(input: XlsStyleInput): XlsStyles {
	const fonts: FontSpec[] = input.fonts.length ? input.fonts : [{ name: 'Arial', size: 10 }];
	const numFmts: Record<number, string> = {};
	for (const [id, code] of input.formats) {
		if (BUILTIN_NUMBER_FORMATS[id] !== code) {
			numFmts[id] = code;
		}
	}
	const fills: FillSpec[] = [
		{ type: 'pattern', patternType: 'none' },
		{ type: 'pattern', patternType: 'gray125' },
	];
	const borders: BorderSpec[] = [{}];
	const fillIds = new Map<string, number>(fills.map((f, i) => [JSON.stringify(f), i]));
	const borderIds = new Map<string, number>([[JSON.stringify({}), 0]]);
	const intern = <T>(list: T[], ids: Map<string, number>, item: T): number => {
		const key = JSON.stringify(item);
		let id = ids.get(key);
		if (id === undefined) {
			id = list.length;
			list.push(item);
			ids.set(key, id);
		}
		return id;
	};
	const fillId = (f: FillSpec): number =>
		f.type === 'pattern' && f.patternType === 'none' ? 0 : intern(fills, fillIds, f);

	const toXf = (xf: RawXf, style: boolean): CellXf => {
		const out: CellXf = {
			numFmtId: xf.numFmtId,
			fontId: fontIndex(xf.font, fonts.length),
			fillId: fillId(xf.fill),
			borderId: intern(borders, borderIds, xf.border),
		};
		if (Object.keys(xf.alignment).length > 0) {
			out.alignment = { ...xf.alignment };
		}
		if (!xf.locked || xf.hidden) {
			out.protection = {};
			if (!xf.locked) {
				out.protection.locked = false;
			}
			if (xf.hidden) {
				out.protection.hidden = true;
			}
		}
		if (xf.quotePrefix) {
			out.quotePrefix = true;
		}
		// Cell XFs flag the groups they override; style XFs flag the groups they leave out.
		const groups: Array<[boolean, ApplyFlag]> = [
			[xf.used.num, 'applyNumberFormat'],
			[xf.used.font, 'applyFont'],
			[xf.used.fill, 'applyFill'],
			[xf.used.border, 'applyBorder'],
			[xf.used.align, 'applyAlignment'],
			[xf.used.prot, 'applyProtection'],
		];
		for (const [flag, key] of groups) {
			if (flag) {
				out[key] = !style;
			}
		}
		return out;
	};

	const styleIndex = new Map<number, number>();
	const cellIndex = new Map<number, number>();
	const cellStyleXfs: CellXf[] = [];
	const cellXfs: CellXf[] = [];
	input.xfs.forEach((xf, i) => {
		if (xf.isStyle) {
			styleIndex.set(i, cellStyleXfs.length);
			cellStyleXfs.push(toXf(xf, true));
		}
	});
	if (cellStyleXfs.length === 0) {
		cellStyleXfs.push({ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 });
	}
	input.xfs.forEach((xf, i) => {
		if (!xf.isStyle) {
			cellIndex.set(i, cellXfs.length);
			cellXfs.push({ ...toXf(xf, false), xfId: styleIndex.get(xf.parent) ?? 0 });
		}
	});
	if (cellXfs.length === 0) {
		cellXfs.push({ ...cellStyleXfs[0], xfId: 0 });
	}

	const cellStyles: NamedCellStyle[] = [];
	const names = new Set<string>();
	for (const s of input.styles) {
		const name = s.builtinId !== undefined ? builtinStyleName(s.builtinId, s.level ?? 0) : s.name;
		const xfId = styleIndex.get(s.xf);
		if (!name || xfId === undefined || names.has(name.toLowerCase())) {
			continue;
		}
		names.add(name.toLowerCase());
		const entry: NamedCellStyle = { name, xfId };
		if (s.builtinId !== undefined) {
			entry.builtinId = s.builtinId;
			if (s.builtinId === 1 || s.builtinId === 2) {
				entry.iLevel = s.level ?? 0;
			}
		}
		cellStyles.push(entry);
	}
	if (!cellStyles.some((s) => s.builtinId === 0)) {
		cellStyles.unshift({
			name: names.has('normal') ? 'Normal 1' : 'Normal',
			xfId: 0,
			builtinId: 0,
		});
	}

	const sheet: StyleSheet = {
		numFmts,
		fonts,
		fills,
		borders,
		cellStyleXfs,
		cellXfs,
		cellStyles,
		dxfs: [],
	};
	if (input.palette?.length) {
		sheet.indexedColors = DEFAULT_INDEXED_COLORS.map(
			(c, i) => `FF${(i >= 8 ? input.palette![i - 8] : undefined) ?? c}`,
		);
	}

	// Cells pointing at a style XF get a cell XF derived from it (created once).
	const fromStyle = new Map<number, number>();
	const result: XlsStyles = {
		sheet,
		badXf: false,
		cellXf(ixfe: number): number {
			const direct = cellIndex.get(ixfe);
			if (direct !== undefined) {
				return direct;
			}
			const style = styleIndex.get(ixfe);
			if (style === undefined) {
				result.badXf = true;
				return 0;
			}
			let id = fromStyle.get(ixfe);
			if (id === undefined) {
				id = cellXfs.length;
				const src = cellStyleXfs[style];
				const xf: CellXf = {
					numFmtId: src.numFmtId,
					fontId: src.fontId,
					fillId: src.fillId,
					borderId: src.borderId,
					xfId: style,
				};
				if (src.alignment) {
					xf.alignment = { ...src.alignment };
				}
				if (src.protection) {
					xf.protection = { ...src.protection };
				}
				cellXfs.push(xf);
				fromStyle.set(ixfe, id);
			}
			return id;
		},
		font(ifnt: number): FontSpec | undefined {
			return ifnt === 4 ? undefined : fonts[ifnt >= 4 ? ifnt - 1 : ifnt];
		},
	};
	return result;
}
