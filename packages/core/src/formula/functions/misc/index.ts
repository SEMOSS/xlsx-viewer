import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import { collectNumbers, toNumber, toText } from '../../coerce';
import type { FunctionDefinition, RefValue, Scalar } from '../../types';
import { isRefValue } from '../../types';
import { bahtText } from './bahttext';
import { evaluateXPath } from './xpath';

/** PERCENTOF(data_subset, data_all): the subset's sum as a share of the total. */
const PERCENTOF: FunctionDefinition = {
	name: 'PERCENTOF',
	minArgs: 2,
	maxArgs: 2,
	params: ['any', 'any'],
	impl: (args, ctx) => {
		const part = collectNumbers([args[0]], ctx);
		const all = collectNumbers([args[1]], ctx);
		if (isCellError(part)) {
			return part;
		}
		if (isCellError(all)) {
			return all;
		}
		const total = all.reduce((a, b) => a + b, 0);
		return total === 0 ? ERRORS.DIV0 : part.reduce((a, b) => a + b, 0) / total;
	},
};

/** RFC 3986 unreserved characters stay as they are; everything else is percent-encoded (UTF-8). */
const ENCODEURL: FunctionDefinition = {
	name: 'ENCODEURL',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	impl: (args) => {
		const text = toText(args[0] as Scalar);
		if (isCellError(text)) {
			return text;
		}
		return encodeURIComponent(text).replace(
			/[!'()*]/g,
			(c) => `%${c.charCodeAt(0).toString(16).toUpperCase()}`,
		);
	},
};

const BAHTTEXT: FunctionDefinition = {
	name: 'BAHTTEXT',
	minArgs: 1,
	maxArgs: 1,
	params: ['value'],
	impl: (args, ctx) => {
		const n = toNumber(args[0] as Scalar, ctx.date1904);
		return isCellError(n) ? n : bahtText(n);
	},
};

/**
 * TRIMRANGE(range, [trim_rows], [trim_cols]): the reference without its
 * blank outer rows / columns. 0 keeps, 1 trims leading, 2 trailing, 3 both
 * (the default). A completely blank range is `#REF!`.
 */
const TRIMRANGE: FunctionDefinition = {
	name: 'TRIMRANGE',
	minArgs: 1,
	maxArgs: 3,
	params: ['ref', 'value', 'value'],
	impl: (args, ctx) => {
		const ref = args[0];
		if (!isRefValue(ref)) {
			return isCellError(ref) ? ref : ERRORS.VALUE;
		}
		const mode = (i: number): number | CellError => {
			const v = args[i];
			if (
				v === undefined ||
				(typeof v === 'object' && v !== null && 'type' in v && v.type === 'missing')
			) {
				return 3;
			}
			const n = toNumber(v as Scalar, ctx.date1904);
			return isCellError(n) ? n : n >= 0 && n <= 3 ? Math.trunc(n) : ERRORS.VALUE;
		};
		const rowsMode = mode(1);
		const colsMode = mode(2);
		if (isCellError(rowsMode)) {
			return rowsMode;
		}
		if (isCellError(colsMode)) {
			return colsMode;
		}
		let { top, left, bottom, right } = ref.range;
		// Only the used part of the sheet can hold values (whole-column references stay cheap).
		const used = ctx.usedRange(ref.sheet);
		if (!used) {
			return ERRORS.REF;
		}
		const scanBottom = Math.min(bottom, used.bottom);
		const scanRight = Math.min(right, used.right);
		let minR = Infinity;
		let maxR = -1;
		let minC = Infinity;
		let maxC = -1;
		for (let r = top; r <= scanBottom; r++) {
			for (let c = left; c <= scanRight; c++) {
				const v = ctx.getCellValue(ref.sheet, r, c);
				if (v !== null && v !== '') {
					minR = Math.min(minR, r);
					maxR = Math.max(maxR, r);
					minC = Math.min(minC, c);
					maxC = Math.max(maxC, c);
				}
			}
		}
		if (maxR < 0) {
			return ERRORS.REF;
		}
		if (rowsMode & 1) {
			top = minR;
		}
		if (rowsMode & 2) {
			bottom = maxR;
		}
		if (colsMode & 1) {
			left = minC;
		}
		if (colsMode & 2) {
			right = maxC;
		}
		const out: RefValue = { kind: 'ref', sheet: ref.sheet, range: { top, left, bottom, right } };
		return out;
	},
};

/**
 * FILTERXML(xml, xpath): the values XPath selects, as a vertical array
 * (numbers stay numbers). Nothing selected, bad XML or an unsupported path
 * is `#VALUE!`.
 */
const FILTERXML: FunctionDefinition = {
	name: 'FILTERXML',
	minArgs: 2,
	maxArgs: 2,
	params: ['value', 'value'],
	impl: (args) => {
		const xml = toText(args[0] as Scalar);
		const path = toText(args[1] as Scalar);
		if (isCellError(xml)) {
			return xml;
		}
		if (isCellError(path)) {
			return path;
		}
		const found = evaluateXPath(xml, path);
		if (!found || found.length === 0) {
			return ERRORS.VALUE;
		}
		const values = found.map(
			(t): Scalar => (t.trim() !== '' && Number.isFinite(Number(t)) ? Number(t) : t),
		);
		return values.length === 1
			? values[0]
			: { kind: 'array', rows: values.length, cols: 1, values: values.map((v) => [v]) };
	},
};

/** Functions that need an online service or an OLAP connection: unavailable offline, as in a blocked Excel. */
function unavailable(
	name: string,
	minArgs: number,
	maxArgs: number,
	error = ERRORS.BLOCKED,
): FunctionDefinition {
	return { name, minArgs, maxArgs, params: ['any'], impl: () => error };
}

export const MISC_FUNCTIONS: FunctionDefinition[] = [
	PERCENTOF,
	ENCODEURL,
	BAHTTEXT,
	TRIMRANGE,
	FILTERXML,
	unavailable('WEBSERVICE', 1, 1),
	unavailable('IMAGE', 1, 5),
	unavailable('STOCKHISTORY', 2, 255),
	unavailable('TRANSLATE', 1, 3),
	unavailable('DETECTLANGUAGE', 1, 1),
	unavailable('RTD', 3, 255),
	unavailable('FIELDVALUE', 2, 2, ERRORS.FIELD),
	unavailable('CUBEMEMBER', 2, 3, ERRORS.NA),
	unavailable('CUBEVALUE', 1, 255, ERRORS.NA),
	unavailable('CUBESET', 2, 5, ERRORS.NA),
	unavailable('CUBESETCOUNT', 1, 1, ERRORS.NA),
	unavailable('CUBERANKEDMEMBER', 3, 4, ERRORS.NA),
	unavailable('CUBEMEMBERPROPERTY', 3, 3, ERRORS.NA),
	unavailable('CUBEKPIMEMBER', 3, 4, ERRORS.NA),
];
