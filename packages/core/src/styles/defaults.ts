import type { StyleSheet } from '../types/style';
import type { Theme } from '../types/theme';

/** The Office 2013-2022 theme (Calibri), used for new workbooks and as a fallback. */
export function createDefaultTheme(): Theme {
	return {
		name: 'Office Theme',
		colors: {
			name: 'Office',
			dk1: '000000',
			lt1: 'FFFFFF',
			dk2: '44546A',
			lt2: 'E7E6E6',
			accent1: '4472C4',
			accent2: 'ED7D31',
			accent3: 'A5A5A5',
			accent4: 'FFC000',
			accent5: '5B9BD5',
			accent6: '70AD47',
			hlink: '0563C1',
			folHlink: '954F72',
		},
		fonts: {
			name: 'Office',
			majorLatin: 'Calibri Light',
			minorLatin: 'Calibri',
		},
	};
}

/** The styles part Excel writes for a blank workbook. */
export function createDefaultStyleSheet(): StyleSheet {
	return {
		numFmts: {},
		fonts: [{ name: 'Calibri', size: 11, color: { theme: 1 }, family: 2, scheme: 'minor' }],
		fills: [
			{ type: 'pattern', patternType: 'none' },
			{ type: 'pattern', patternType: 'gray125' },
		],
		borders: [{}],
		cellStyleXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }],
		cellXfs: [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0, xfId: 0 }],
		cellStyles: [{ name: 'Normal', xfId: 0, builtinId: 0 }],
		dxfs: [],
		defaultTableStyle: 'TableStyleMedium2',
		defaultPivotStyle: 'PivotStyleLight16',
	};
}

/** Default row height in points for the default 11pt body font. */
export const DEFAULT_ROW_HEIGHT_PT = 15;
/** Default stored column width (8.43 visible characters plus padding) for Calibri 11. */
export const DEFAULT_COL_WIDTH_CHARS = 9.140625;
/** Maximum digit width in pixels of Calibri 11 at 96 DPI. */
export const DEFAULT_MAX_DIGIT_WIDTH_PX = 7;
