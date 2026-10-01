/** `<dataValidations>` plus the Excel 2010 `x14:dataValidations` extension. */
import { formatSqref } from '../address/range';
import type { DataValidation } from '../types/data-validation';
import { el, textEl } from '../xml';
import { NS } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { fileFormula } from './formulas';
import { flag } from './xml-utils';

export const DV_EXT_URI = '{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}';

/** Excel limits prompt / error texts; longer ones make it repair the file. */
const MAX_TITLE = 32;
const MAX_MESSAGE = 255;

function clip(text: string | undefined, max: number): string | undefined {
	return text !== undefined && text.length > max ? text.slice(0, max) : text;
}

const OPERATOR_TYPES = new Set(['whole', 'decimal', 'date', 'time', 'textLength']);

function dvAttrs(dv: DataValidation): Record<string, string | number | boolean | undefined> {
	return {
		type: dv.type && dv.type !== 'none' ? dv.type : undefined,
		errorStyle: dv.errorStyle && dv.errorStyle !== 'stop' ? dv.errorStyle : undefined,
		operator:
			OPERATOR_TYPES.has(dv.type) && dv.operator && dv.operator !== 'between'
				? dv.operator
				: undefined,
		allowBlank: flag(dv.allowBlank),
		showDropDown: dv.type === 'list' && dv.inCellDropdown === false ? 1 : undefined,
		showInputMessage: flag(dv.showInputMessage),
		showErrorMessage: flag(dv.showErrorMessage),
		errorTitle: clip(dv.errorTitle, MAX_TITLE),
		error: clip(dv.error, MAX_MESSAGE),
		promptTitle: clip(dv.promptTitle, MAX_TITLE),
		prompt: clip(dv.prompt, MAX_MESSAGE),
	};
}

export interface DvOutput {
	element: XmlElement | null;
	ext: XmlElement | null;
}

/** Serialize a sheet's data validations. */
export function buildDataValidations(list: readonly DataValidation[]): DvOutput {
	const std: XmlElement[] = [];
	const ext: XmlElement[] = [];
	for (const dv of list) {
		if (!dv.ranges.length) {
			continue;
		}
		const sqref = formatSqref(dv.ranges);
		const f1 =
			dv.formula1 !== undefined && dv.formula1 !== '' ? fileFormula(dv.formula1) : undefined;
		const f2 =
			dv.formula2 !== undefined && dv.formula2 !== '' ? fileFormula(dv.formula2) : undefined;
		if (dv.fromExtension) {
			ext.push(
				el('x14:dataValidation', dvAttrs(dv), [
					f1 !== undefined ? el('x14:formula1', {}, [textEl('xm:f', f1)]) : undefined,
					f2 !== undefined ? el('x14:formula2', {}, [textEl('xm:f', f2)]) : undefined,
					textEl('xm:sqref', sqref),
				]),
			);
		} else {
			std.push(
				el('dataValidation', { ...dvAttrs(dv), sqref }, [
					f1 !== undefined ? textEl('formula1', f1) : undefined,
					f2 !== undefined ? textEl('formula2', f2) : undefined,
				]),
			);
		}
	}
	return {
		element: std.length ? el('dataValidations', { count: std.length }, std) : null,
		ext: ext.length
			? el('ext', { uri: DV_EXT_URI, 'xmlns:x14': NS.x14 }, [
					el('x14:dataValidations', { count: ext.length, 'xmlns:xm': NS.xm }, ext),
				])
			: null,
	};
}
