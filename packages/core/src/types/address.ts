/**
 * Cell and range coordinates.
 *
 * All coordinates in the data model are **0-based**: row 0 is Excel row 1 and
 * column 0 is column A. Ranges are inclusive on both ends and normalised so
 * that `top <= bottom` and `left <= right`.
 */

/** A single cell position (0-based). */
export interface CellAddress {
	row: number;
	col: number;
}

/** An inclusive rectangular block of cells (0-based, normalised). */
export interface RangeAddress {
	top: number;
	left: number;
	bottom: number;
	right: number;
}
