/**
 * Page setup, print options, sheet protection and the calculation settings
 * every BIFF8 worksheet repeats (the workbook takes them from the first).
 */
import type {
	HeaderFooter,
	PageMargins,
	PageSetup,
	PrintOptions,
	SheetProtection,
} from '../../types/sheet';
import type { CalcProperties } from '../../types/workbook';
import { BiffCursor } from './biff-cursor';
import { RT } from './biff-opcodes';
import type { BiffRecord } from './biff-records';
import type { SheetContext } from './sheet-context';

/** Excel 97-2003 default margins in inches. */
const DEFAULT_MARGINS: PageMargins = {
	left: 0.75,
	right: 0.75,
	top: 1,
	bottom: 1,
	header: 0.5,
	footer: 0.5,
};

const MARGIN_KEYS: Readonly<Record<number, keyof PageMargins>> = {
	[RT.LEFTMARGIN]: 'left',
	[RT.RIGHTMARGIN]: 'right',
	[RT.TOPMARGIN]: 'top',
	[RT.BOTTOMMARGIN]: 'bottom',
};

function readSetup(cur: BiffCursor, margins: Partial<PageMargins>): PageSetup {
	const [paper, scale, start, fitWidth, fitHeight, flags, hRes, vRes] = [
		cur.u16(),
		cur.u16(),
		cur.i16(),
		cur.u16(),
		cur.u16(),
		cur.u16(),
		cur.u16(),
		cur.u16(),
	];
	if (cur.remaining >= 16) {
		margins.header = cur.f64();
		margins.footer = cur.f64();
	}
	const setup: PageSetup = {};
	if (!(flags & 0x04)) {
		if (paper) {
			setup.paperSize = paper;
		}
		if (scale && scale !== 100) {
			setup.scale = scale;
		}
		setup.orientation = flags & 0x40 ? 'default' : flags & 0x02 ? 'portrait' : 'landscape';
		if (hRes) {
			setup.horizontalDpi = hRes;
		}
		if (vRes) {
			setup.verticalDpi = vRes;
		}
	}
	if (fitWidth !== 1) {
		setup.fitToWidth = fitWidth;
	}
	if (fitHeight !== 1) {
		setup.fitToHeight = fitHeight;
	}
	if (flags & 0x80) {
		setup.firstPageNumber = start;
		setup.useFirstPageNumber = true;
	}
	if (flags & 0x01) {
		setup.pageOrder = 'overThenDown';
	}
	if (flags & 0x08) {
		setup.blackAndWhite = true;
	}
	if (flags & 0x10) {
		setup.draft = true;
	}
	return setup;
}

/** Record types handled by {@link SettingsReader.read}. */
export const SETTINGS_RECORDS: readonly number[] = [
	RT.LEFTMARGIN,
	RT.RIGHTMARGIN,
	RT.TOPMARGIN,
	RT.BOTTOMMARGIN,
	RT.SETUP,
	RT.HEADER,
	RT.FOOTER,
	RT.HCENTER,
	RT.VCENTER,
	RT.PRINTGRIDLINES,
	RT.PRINTHEADERS,
	RT.PROTECT,
	RT.OBJPROTECT,
	RT.SCENPROTECT,
	RT.PASSWORD,
	RT.CALCMODE,
	RT.CALCCOUNT,
	RT.ITERATION,
	RT.DELTA,
	RT.REFMODE,
	RT.AUTOFILTERINFO,
];

export class SettingsReader {
	readonly calc: CalcProperties = {};
	/** The sheet has an AutoFilter (its range is the sheet's `_FilterDatabase` name). */
	autoFilter = false;
	private readonly margins: Partial<PageMargins> = {};
	private setup: PageSetup | undefined;
	private readonly print: PrintOptions = {};
	private readonly headerFooter: HeaderFooter = {};
	private readonly protection: Partial<SheetProtection> = {};

	constructor(private readonly ctx: SheetContext) {}

	/** Handle a settings record; returns `false` for other record types. */
	read(record: BiffRecord): boolean {
		const cur = new BiffCursor(record.data, record.breaks);
		const margin = MARGIN_KEYS[record.type];
		if (margin) {
			this.margins[margin] = cur.f64();
			return true;
		}
		const flag = (): boolean => cur.remaining >= 2 && cur.u16() !== 0;
		switch (record.type) {
			case RT.SETUP:
				this.setup = readSetup(cur, this.margins);
				return true;
			case RT.HEADER:
			case RT.FOOTER: {
				const text = cur.remaining >= 3 ? cur.unicodeString() : '';
				if (text) {
					this.headerFooter[record.type === RT.HEADER ? 'oddHeader' : 'oddFooter'] = text;
				}
				return true;
			}
			case RT.HCENTER:
				if (flag()) {
					this.print.horizontalCentered = true;
				}
				return true;
			case RT.VCENTER:
				if (flag()) {
					this.print.verticalCentered = true;
				}
				return true;
			case RT.PRINTGRIDLINES:
				if (flag()) {
					this.print.gridLines = true;
				}
				return true;
			case RT.PRINTHEADERS:
				if (flag()) {
					this.print.headings = true;
				}
				return true;
			case RT.PROTECT:
				this.protection.sheet = flag();
				return true;
			case RT.OBJPROTECT:
				if (flag()) {
					this.protection.objects = true;
				}
				return true;
			case RT.SCENPROTECT:
				if (flag()) {
					this.protection.scenarios = true;
				}
				return true;
			case RT.PASSWORD: {
				const hash = cur.u16();
				if (hash) {
					this.protection.password = hash.toString(16).toUpperCase().padStart(4, '0');
				}
				return true;
			}
			case RT.CALCMODE: {
				const mode = cur.u16();
				if (mode !== 1) {
					this.calc.calcMode = mode === 0 ? 'manual' : 'autoNoTable';
				}
				return true;
			}
			case RT.CALCCOUNT: {
				const count = cur.u16();
				if (count !== 100) {
					this.calc.iterateCount = count;
				}
				return true;
			}
			case RT.ITERATION:
				if (flag()) {
					this.calc.iterate = true;
				}
				return true;
			case RT.DELTA: {
				const delta = cur.f64();
				if (delta !== 0.001) {
					this.calc.iterateDelta = delta;
				}
				return true;
			}
			case RT.REFMODE:
				if (cur.u16() === 0) {
					this.calc.refMode = 'R1C1';
				}
				return true;
			case RT.AUTOFILTERINFO:
				this.autoFilter = true;
				return true;
			default:
				return false;
		}
	}

	finish(): void {
		const sheet = this.ctx.sheet;
		if (Object.keys(this.margins).length > 0) {
			sheet.pageMargins = { ...DEFAULT_MARGINS, ...this.margins };
		}
		if (this.setup && Object.keys(this.setup).length > 0) {
			sheet.pageSetup = this.setup;
		}
		if (Object.keys(this.print).length > 0) {
			sheet.printOptions = this.print;
		}
		if (Object.keys(this.headerFooter).length > 0) {
			sheet.headerFooter = this.headerFooter;
		}
		if (this.protection.sheet) {
			const p: SheetProtection = { sheet: true };
			if (this.protection.objects) {
				p.objects = true;
			}
			if (this.protection.scenarios) {
				p.scenarios = true;
			}
			if (this.protection.password) {
				p.password = this.protection.password;
			}
			sheet.protection = p;
		}
	}
}
