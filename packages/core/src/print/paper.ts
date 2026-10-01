import type { PageMargins, PageSetup } from '../types/sheet';

/**
 * OOXML paper sizes (`pageSetup/@paperSize`, ECMA-376 18.3.1.63) in points.
 * Codes 12 and 13 are the JIS B4 / B5 sheets Windows printers use for them.
 */

export interface PaperSize {
	code: number;
	name: string;
	/** Portrait width and height in points. */
	width: number;
	height: number;
}

const IN = 72;
const MM = 72 / 25.4;

function inch(code: number, name: string, w: number, h: number): PaperSize {
	return { code, name, width: w * IN, height: h * IN };
}

function mm(code: number, name: string, w: number, h: number): PaperSize {
	return { code, name, width: w * MM, height: h * MM };
}

/** Paper sizes by code, in the order the Size menu lists them. */
export const PAPER_SIZES: readonly PaperSize[] = [
	inch(1, 'Letter', 8.5, 11),
	inch(5, 'Legal', 8.5, 14),
	inch(7, 'Executive', 7.25, 10.5),
	mm(9, 'A4', 210, 297),
	mm(11, 'A5', 148, 210),
	mm(8, 'A3', 297, 420),
	mm(66, 'A2', 420, 594),
	mm(13, 'B5 (JIS)', 182, 257),
	mm(12, 'B4 (JIS)', 257, 364),
	inch(3, 'Tabloid', 11, 17),
	inch(4, 'Ledger', 17, 11),
	inch(6, 'Statement', 5.5, 8.5),
	inch(14, 'Folio', 8.5, 13),
	mm(15, 'Quarto', 215, 275),
	inch(16, '10x14', 10, 14),
	inch(17, '11x17', 11, 17),
	inch(18, 'Note', 8.5, 11),
	inch(2, 'Letter Small', 8.5, 11),
	mm(10, 'A4 Small', 210, 297),
	inch(19, 'Envelope #9', 3.875, 8.875),
	inch(20, 'Envelope #10', 4.125, 9.5),
	inch(21, 'Envelope #11', 4.5, 10.375),
	inch(22, 'Envelope #12', 4.75, 11),
	inch(23, 'Envelope #14', 5, 11.5),
	inch(24, 'C Size Sheet', 17, 22),
	inch(25, 'D Size Sheet', 22, 34),
	inch(26, 'E Size Sheet', 34, 44),
	mm(27, 'Envelope DL', 110, 220),
	mm(28, 'Envelope C5', 162, 229),
	mm(29, 'Envelope C3', 324, 458),
	mm(30, 'Envelope C4', 229, 324),
	mm(31, 'Envelope C6', 114, 162),
	mm(32, 'Envelope C65', 114, 229),
	mm(33, 'Envelope B4', 250, 353),
	mm(34, 'Envelope B5', 176, 250),
	mm(35, 'Envelope B6', 176, 125),
	mm(36, 'Envelope Italy', 110, 230),
	inch(37, 'Envelope Monarch', 3.875, 7.5),
	inch(38, '6 3/4 Envelope', 3.625, 6.5),
	inch(39, 'US Std Fanfold', 14.875, 11),
	inch(40, 'German Std Fanfold', 8.5, 12),
	inch(41, 'German Legal Fanfold', 8.5, 13),
	mm(42, 'B4 (ISO)', 250, 353),
	mm(43, 'Japanese Double Postcard', 200, 148),
	inch(44, '9x11', 9, 11),
	inch(45, '10x11', 10, 11),
	inch(46, '15x11', 15, 11),
	mm(47, 'Envelope Invite', 220, 220),
	inch(50, 'Letter Extra', 9.275, 12),
	inch(51, 'Legal Extra', 9.275, 15),
	inch(52, 'Tabloid Extra', 11.69, 18),
	mm(53, 'A4 Extra', 236, 322),
	inch(54, 'Letter Transverse', 8.275, 11),
	mm(55, 'A4 Transverse', 210, 297),
	inch(56, 'Letter Extra Transverse', 9.275, 12),
	mm(57, 'Super A', 227, 356),
	mm(58, 'Super B', 305, 487),
	inch(59, 'Letter Plus', 8.5, 12.69),
	mm(60, 'A4 Plus', 210, 330),
	mm(61, 'A5 Transverse', 148, 210),
	mm(62, 'B5 (JIS) Transverse', 182, 257),
	mm(63, 'A3 Extra', 322, 445),
	mm(64, 'A5 Extra', 174, 235),
	mm(65, 'B5 (ISO) Extra', 201, 276),
	mm(67, 'A3 Transverse', 297, 420),
	mm(68, 'A3 Extra Transverse', 322, 445),
];

const BY_CODE = new Map(PAPER_SIZES.map((p) => [p.code, p]));

/** The schema default paper (Letter) when the sheet does not say. */
export const DEFAULT_PAPER_SIZE = 1;

/** The sizes offered first in Excel's Size menu. */
export const COMMON_PAPER_SIZES: readonly number[] = [1, 5, 7, 9, 11, 8, 13, 3, 20, 27, 28];

/** Paper size for a code (Letter for unknown codes). */
export function paperSize(code: number | undefined): PaperSize {
	return BY_CODE.get(code ?? DEFAULT_PAPER_SIZE) ?? BY_CODE.get(DEFAULT_PAPER_SIZE)!;
}

/** Paper width and height in points for a page setup (orientation applied). */
export function paperDimensions(setup: PageSetup | undefined): {
	width: number;
	height: number;
	landscape: boolean;
} {
	const p = paperSize(setup?.paperSize);
	const landscape = setup?.orientation === 'landscape';
	const [w, h] = [Math.min(p.width, p.height), Math.max(p.width, p.height)];
	return landscape ? { width: h, height: w, landscape } : { width: w, height: h, landscape };
}

/** A short description such as `8.5" x 11"` or `21 cm x 29.7 cm`. */
export function paperSizeDetail(p: PaperSize, metric = false): string {
	if (metric) {
		const cm = (pt: number) => `${Math.round((pt / MM) * 10) / 100} cm`;
		return `${cm(p.width)} x ${cm(p.height)}`;
	}
	const inches = (pt: number) => `${Math.round((pt / IN) * 100) / 100}"`;
	return `${inches(p.width)} x ${inches(p.height)}`;
}

/** Margin presets of Page Layout > Margins, in inches. */
export type MarginPreset = 'normal' | 'wide' | 'narrow';

export const MARGIN_PRESETS: Readonly<Record<MarginPreset, PageMargins>> = {
	normal: { top: 0.75, bottom: 0.75, left: 0.7, right: 0.7, header: 0.3, footer: 0.3 },
	wide: { top: 1, bottom: 1, left: 1, right: 1, header: 0.5, footer: 0.5 },
	narrow: { top: 0.75, bottom: 0.75, left: 0.25, right: 0.25, header: 0.3, footer: 0.3 },
};

/** Excel's default margins (the Normal preset). */
export const DEFAULT_MARGINS: PageMargins = MARGIN_PRESETS.normal;

/** The preset the margins match, if any. */
export function marginPresetOf(m: PageMargins | undefined): MarginPreset | undefined {
	const margins = m ?? DEFAULT_MARGINS;
	const same = (a: number, b: number) => Math.abs(a - b) < 0.001;
	for (const [name, p] of Object.entries(MARGIN_PRESETS) as Array<[MarginPreset, PageMargins]>) {
		if (
			same(p.top, margins.top) &&
			same(p.bottom, margins.bottom) &&
			same(p.left, margins.left) &&
			same(p.right, margins.right) &&
			same(p.header, margins.header) &&
			same(p.footer, margins.footer)
		) {
			return name;
		}
	}
	return undefined;
}
