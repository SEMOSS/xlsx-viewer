import type { Rgb } from '../color/color-math';
import type { ColorContext } from '../color/resolve-color';
import type { RangeAddress } from '../types/address';
import type { CellValue } from '../types/cell';
import type { CfRule } from '../types/conditional-format';
import type { Worksheet } from '../types/sheet';
import type { DifferentialFormat } from '../types/style';
import type { Workbook } from '../types/workbook';
import type { DataBarStyle } from './data-bar';
import type { CellReader, FormulaEvaluator, RelativeFormula } from './formula-eval';
import { anyRangeContains, boundingBox } from './range-cells';
import { RangeStatistics } from './statistics';

/** Per-sheet environment shared by every rule evaluation. */
export interface RuleEnv {
	readonly workbook: Workbook;
	readonly sheet: Worksheet;
	readonly sheetIndex: number;
	readonly evaluate: FormulaEvaluator;
	readonly read: CellReader;
	readonly today: () => number;
	readonly colorCtx: ColorContext;
}

/**
 * One `<conditionalFormatting>` block: its (sanitised, non-empty) ranges,
 * bounding box, formula anchor and lazily computed range statistics, shared
 * by all rules of the block.
 */
export class CfBlock {
	readonly bbox: RangeAddress;
	private stats: RangeStatistics | undefined;

	constructor(
		private readonly sheet: Worksheet,
		readonly ranges: readonly RangeAddress[],
	) {
		this.bbox = boundingBox(ranges) ?? { top: 0, left: 0, bottom: -1, right: -1 };
	}

	/** Formulas are relative to the top-left cell of the first range. */
	get anchorRow(): number {
		return this.ranges[0]?.top ?? 0;
	}

	get anchorCol(): number {
		return this.ranges[0]?.left ?? 0;
	}

	contains(row: number, col: number): boolean {
		const b = this.bbox;
		if (row < b.top || row > b.bottom || col < b.left || col > b.right) {
			return false;
		}
		return this.ranges.length === 1 || anyRangeContains(this.ranges, row, col);
	}

	get statistics(): RangeStatistics {
		this.stats ??= new RangeStatistics(this.sheet, this.ranges);
		return this.stats;
	}
}

/** Lazily prepared, rule-specific evaluation state (dropped on invalidation). */
export type RuleState =
	| { kind: 'invalid' }
	| { kind: 'formula'; formulas: RelativeFormula[] }
	| { kind: 'match'; test: (value: CellValue) => boolean }
	| { kind: 'colorScale'; thresholds: number[]; colors: Rgb[] }
	| { kind: 'dataBar'; lo: number; hi: number; style: DataBarStyle }
	| { kind: 'iconSet'; thresholds: number[]; gte: boolean[]; count: number };

export const INVALID_RULE: RuleState = Object.freeze({ kind: 'invalid' });

export interface CompiledRule {
	readonly rule: CfRule;
	readonly block: CfBlock;
	/** Resolved differential format (inline `dxf` or `styles.dxfs[dxfId]`). */
	readonly dxf: DifferentialFormat | undefined;
	/** Document order, the tie-breaker between equal priorities. */
	readonly order: number;
	state?: RuleState;
}
