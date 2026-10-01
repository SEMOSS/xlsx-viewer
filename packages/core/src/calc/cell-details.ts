import { cellKey } from '../address/cell-address';
import type { CellInfoProvider } from '../formula/functions/information/cell';
import { getColInfo } from '../sheet/worksheet';
import { storedWidthToDisplayChars } from '../styles/column-width';
import { type ResolvedCellStyle, StyleResolver } from '../styles/resolve-style';
import type { Workbook } from '../types/workbook';

/**
 * Formatting and layout details for CELL(): number format, alignment and
 * protection from the cell's resolved style (falling back to the row and
 * column default styles), column widths from the column records.
 */
export class CellDetails implements CellInfoProvider {
	private resolver: StyleResolver | null = null;
	private resolverSize = -1;

	constructor(private readonly workbook: Workbook) {}

	/** Forget the cached style resolver (after the stylesheet was replaced). */
	reset(): void {
		this.resolver = null;
	}

	private styleOf(sheet: number, row: number, col: number): number {
		const ws = this.workbook.sheets[sheet];
		if (ws === undefined) {
			return 0;
		}
		const cell = ws.cells.get(cellKey(row, col));
		if (cell?.style !== undefined) {
			return cell.style;
		}
		const info = ws.rows.get(row);
		if (info?.customFormat && info.style !== undefined) {
			return info.style;
		}
		return getColInfo(ws, col)?.style ?? 0;
	}

	private resolved(sheet: number, row: number, col: number): ResolvedCellStyle {
		const styles = this.workbook.styles;
		if (
			this.resolver === null ||
			this.resolver.styles !== styles ||
			this.resolverSize !== styles.cellXfs.length
		) {
			this.resolver = new StyleResolver(styles, this.workbook.theme);
			this.resolverSize = styles.cellXfs.length;
		}
		return this.resolver.resolve(this.styleOf(sheet, row, col));
	}

	getNumberFormat(sheet: number, row: number, col: number): string | undefined {
		const code = this.resolved(sheet, row, col).numFmt;
		return /^general$/i.test(code) ? undefined : code;
	}

	getHorizontalAlignment(sheet: number, row: number, col: number): string | undefined {
		return this.resolved(sheet, row, col).alignment.horizontal;
	}

	isCellLocked(sheet: number, row: number, col: number): boolean {
		return this.resolved(sheet, row, col).protection.locked;
	}

	getColumnWidth(sheet: number, col: number): { width: number; isDefault: boolean } | undefined {
		const ws = this.workbook.sheets[sheet];
		if (ws === undefined) {
			return undefined;
		}
		const info = getColInfo(ws, col);
		if (info?.hidden) {
			return { width: 0, isDefault: false };
		}
		if (info?.width !== undefined) {
			return { width: storedWidthToDisplayChars(info.width), isDefault: false };
		}
		const width =
			ws.defaultColWidth !== undefined
				? storedWidthToDisplayChars(ws.defaultColWidth)
				: (ws.baseColWidth ?? 8) + 0.43;
		return { width, isDefault: true };
	}

	getWorkbookName(): string | undefined {
		return undefined;
	}
}
