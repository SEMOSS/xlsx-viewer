import type { XmlElement } from '../xml/types';
import type { CellAddress, RangeAddress } from './address';
import type { Cell } from './cell';
import type { Note, ThreadedComment } from './comment';
import type { ConditionalFormat } from './conditional-format';
import type { DataValidation } from './data-validation';
import type { DrawingObject } from './drawing';
import type { AutoFilter, SortState } from './filter';
import type { ColorSpec } from './style';
import type { Table } from './table';

/** Per-row formatting and layout (`<row>` attributes). Keys of `Worksheet.rows` are 0-based rows. */
export interface RowInfo {
	/** Height in points (`ht`). */
	height?: number;
	customHeight?: boolean;
	hidden?: boolean;
	/** Row default style (cellXfs index); only applied when `customFormat` is set. */
	style?: number;
	customFormat?: boolean;
	outlineLevel?: number;
	collapsed?: boolean;
	thickTop?: boolean;
	thickBot?: boolean;
}

/** A `<col>` record covering columns `min..max` (0-based, inclusive). */
export interface ColInfo {
	min: number;
	max: number;
	/** Width in character units, as stored in the file (includes padding). */
	width?: number;
	customWidth?: boolean;
	hidden?: boolean;
	/** Column default style (cellXfs index). */
	style?: number;
	outlineLevel?: number;
	collapsed?: boolean;
	bestFit?: boolean;
}

export type PaneName = 'bottomRight' | 'topRight' | 'bottomLeft' | 'topLeft';

export interface SheetPane {
	/** Frozen: number of frozen columns. Split: horizontal split position in 1/20 pt. */
	xSplit: number;
	/** Frozen: number of frozen rows. Split: vertical split position in 1/20 pt. */
	ySplit: number;
	/** Top-left visible cell of the bottom-right pane. */
	topLeftCell?: CellAddress;
	activePane?: PaneName;
	state: 'frozen' | 'frozenSplit' | 'split';
}

export interface SheetSelection {
	activeCell: CellAddress;
	ranges: RangeAddress[];
	pane?: PaneName;
}

/** The first `<sheetView>` of a worksheet. */
export interface SheetView {
	showGridLines: boolean;
	showRowColHeaders: boolean;
	showZeros: boolean;
	showFormulas: boolean;
	rightToLeft: boolean;
	/** Zoom in percent (10-400). */
	zoomScale: number;
	tabSelected?: boolean;
	view?: 'normal' | 'pageBreakPreview' | 'pageLayout';
	/** Top-left visible cell (scroll position). */
	topLeftCell?: CellAddress;
	pane?: SheetPane;
	selection?: SheetSelection;
	/** Gridline color, when overridden. */
	gridColor?: ColorSpec;
}

export interface Hyperlink {
	ref: RangeAddress;
	/** External target (URL, mailto:, file path) from the sheet relationships. */
	target?: string;
	/** Internal location, e.g. `Sheet2!A1` or a defined name. */
	location?: string;
	display?: string;
	tooltip?: string;
}

export interface SheetProtection {
	/** Legacy password hash (hex), when present. */
	password?: string;
	algorithmName?: string;
	hashValue?: string;
	saltValue?: string;
	spinCount?: number;
	sheet: boolean;
	objects?: boolean;
	scenarios?: boolean;
	formatCells?: boolean;
	formatColumns?: boolean;
	formatRows?: boolean;
	insertColumns?: boolean;
	insertRows?: boolean;
	insertHyperlinks?: boolean;
	deleteColumns?: boolean;
	deleteRows?: boolean;
	selectLockedCells?: boolean;
	sort?: boolean;
	autoFilter?: boolean;
	pivotTables?: boolean;
	selectUnlockedCells?: boolean;
}

export interface PageMargins {
	/** All margins in inches. */
	left: number;
	right: number;
	top: number;
	bottom: number;
	header: number;
	footer: number;
}

export interface PageSetup {
	paperSize?: number;
	orientation?: 'default' | 'portrait' | 'landscape';
	scale?: number;
	/**
	 * Fit to `fitToWidth` x `fitToHeight` pages instead of using `scale`
	 * (`sheetPr/pageSetUpPr/@fitToPage`). `undefined` keeps the file's value.
	 */
	fitToPage?: boolean;
	/** Pages wide when fitting (default 1; 0 = automatic). */
	fitToWidth?: number;
	/** Pages tall when fitting (default 1; 0 = automatic). */
	fitToHeight?: number;
	firstPageNumber?: number;
	useFirstPageNumber?: boolean;
	pageOrder?: 'downThenOver' | 'overThenDown';
	blackAndWhite?: boolean;
	draft?: boolean;
	horizontalDpi?: number;
	verticalDpi?: number;
}

export interface PrintOptions {
	gridLines?: boolean;
	headings?: boolean;
	horizontalCentered?: boolean;
	verticalCentered?: boolean;
}

export interface HeaderFooter {
	oddHeader?: string;
	oddFooter?: string;
	evenHeader?: string;
	evenFooter?: string;
	firstHeader?: string;
	firstFooter?: string;
	differentOddEven?: boolean;
	differentFirst?: boolean;
	/** Header / footer text scales with the document (default true). */
	scaleWithDoc?: boolean;
	/** Left / right sections align with the page margins (default true). */
	alignWithMargins?: boolean;
}

export interface Sparkline {
	/** Source data reference, e.g. `Sheet1!A2:F2`. */
	formula: string;
	/** Cell the sparkline is drawn in. */
	location: CellAddress;
}

export interface SparklineGroup {
	type: 'line' | 'column' | 'stacked';
	sparklines: Sparkline[];
	seriesColor?: ColorSpec;
	negativeColor?: ColorSpec;
	markersColor?: ColorSpec;
	highColor?: ColorSpec;
	lowColor?: ColorSpec;
	firstColor?: ColorSpec;
	lastColor?: ColorSpec;
	markers?: boolean;
	high?: boolean;
	low?: boolean;
	first?: boolean;
	last?: boolean;
	negative?: boolean;
	/** Line weight in points. */
	lineWeight?: number;
	displayEmptyCellsAs?: 'gap' | 'zero' | 'span';
	minAxisType?: 'individual' | 'group' | 'custom';
	maxAxisType?: 'individual' | 'group' | 'custom';
	manualMin?: number;
	manualMax?: number;
	rightToLeft?: boolean;
}

export interface Relationship {
	id: string;
	type: string;
	target: string;
	targetMode?: 'External' | 'Internal';
}

/** Round-trip bookkeeping for a sheet loaded from a file. */
export interface SheetSource {
	/** Package path, e.g. `xl/worksheets/sheet1.xml`. */
	partPath: string;
	/** Relationships of the sheet part. */
	rels: Relationship[];
	/**
	 * The original root element with the heavy children the model fully owns
	 * (`sheetData`) emptied. The writer uses it to keep element order and
	 * any markup the model does not understand.
	 */
	root?: XmlElement;
	/** Package paths of related parts (drawing, comments, VML, tables) as loaded. */
	drawingPath?: string;
	commentsPath?: string;
	vmlPath?: string;
	threadedCommentsPath?: string;
}

export type SheetKind = 'worksheet' | 'chartsheet' | 'dialogsheet' | 'macrosheet';

export interface Worksheet {
	/** Stable internal id (survives renames and reordering); not written to the file. */
	id: string;
	name: string;
	/** OOXML `sheetId` from workbook.xml. */
	sheetId: number;
	kind: SheetKind;
	state: 'visible' | 'hidden' | 'veryHidden';
	/** Cells keyed by `cellKey(row, col)`. */
	cells: Map<number, Cell>;
	/** Row records keyed by 0-based row index. */
	rows: Map<number, RowInfo>;
	/** Sorted, non-overlapping column records. */
	cols: ColInfo[];
	/** Default row height in points (`sheetFormatPr/@defaultRowHeight`), typically 15. */
	defaultRowHeight: number;
	/** Default column width in character units including padding (`defaultColWidth`). */
	defaultColWidth?: number;
	/** Base column width in characters (`baseColWidth`), default 8. */
	baseColWidth?: number;
	/** Rows are hidden by default (`zeroHeight`). */
	zeroHeight?: boolean;
	outlineLevelRow?: number;
	outlineLevelCol?: number;
	merges: RangeAddress[];
	view: SheetView;
	tabColor?: ColorSpec;
	hyperlinks: Hyperlink[];
	conditionalFormats: ConditionalFormat[];
	dataValidations: DataValidation[];
	autoFilter?: AutoFilter;
	sortState?: SortState;
	tables: Table[];
	drawings: DrawingObject[];
	notes: Note[];
	threadedComments: ThreadedComment[];
	sparklineGroups: SparklineGroup[];
	protection?: SheetProtection;
	pageMargins?: PageMargins;
	pageSetup?: PageSetup;
	printOptions?: PrintOptions;
	headerFooter?: HeaderFooter;
	/**
	 * Manual page breaks above these 0-based rows (`<rowBreaks><brk id>`; the
	 * OOXML id is the 0-based index of the first row of the next page).
	 * Sorted and unique. `undefined` keeps the file's element as it was.
	 */
	rowBreaks?: number[];
	/** Manual page breaks left of these 0-based columns (`<colBreaks>`). */
	colBreaks?: number[];
	/** Summary rows below detail (outline). Default true. */
	summaryBelow?: boolean;
	/** Summary columns right of detail (outline). Default true. */
	summaryRight?: boolean;
	/** Round-trip data; absent for sheets created in memory. */
	source?: SheetSource;
	/**
	 * Change tracking for save. A sheet with `dirty === false` is written back
	 * byte-for-byte from the source package; anything else is regenerated.
	 */
	dirty?: boolean;
	/** Drawing objects changed (the drawing part must be regenerated). */
	drawingsDirty?: boolean;
	/** Notes / threaded comments changed. */
	commentsDirty?: boolean;
}
