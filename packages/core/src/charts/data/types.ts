/** Live chart data, resolved from worksheet cells (or the chart's cached values). */

export interface ChartSeriesData {
	/** Display name: the referenced cell text, a literal, or `Series N`. */
	name: string;
	/** Category labels as displayed (cell number formats applied), one per point. */
	categories: string[];
	/**
	 * Outer category levels for multi-level category axes, outer level first.
	 * Each level is aligned with `categories`; an empty string continues the
	 * previous group.
	 */
	categoryLevels?: string[][];
	/** Numeric category / X values (scatter, bubble, date axes); `NaN` when not a number. */
	xValues: number[];
	/**
	 * Y values. `null` is a gap (blank cell shown as a gap); `NaN` is a missing
	 * point the line connects over (`#N/A`, or blanks with `displayBlanksAs="span"`).
	 */
	values: Array<number | null>;
	/** Bubble sizes (bubble charts). */
	bubbleSizes?: Array<number | null>;
	/** Number format of the values (source cells, else the cache `formatCode`). */
	valueFormat: string;
	/** Number format of numeric categories / X values. */
	categoryFormat: string;
	/** Every non-blank category is a number (or date serial). */
	categoriesNumeric: boolean;
	/** Values were read from worksheet cells (`false`: cached values from the file). */
	live: boolean;
	/** The source row / column is hidden and the chart plots visible cells only. */
	hidden?: boolean;
}

export interface ChartPlotData {
	/** Parallel to `ChartPlot.series`. */
	series: ChartSeriesData[];
}

export interface ChartData {
	/** Parallel to `ChartModel.plots`. */
	plots: ChartPlotData[];
	/** The workbook uses the 1904 date system (affects date formatting). */
	date1904?: boolean;
	/** Live text of titles linked to cells: the chart title and axis titles by axis id. */
	titles?: { chart?: string; axes?: Record<number, string> };
}
