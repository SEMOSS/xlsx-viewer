import type { XmlElement } from '../xml/types';
import type { ChartModel } from './chart';

/**
 * DrawingML (`xl/drawings/drawingN.xml`) objects floating over a worksheet:
 * pictures, charts, shapes, text boxes and groups.
 *
 * Positions use EMU (English Metric Units): 914400 per inch, 12700 per point,
 * 9525 per pixel at 96 DPI.
 */

/** A DrawingML color with optional transforms (`lumMod`, `tint`, `alpha`, ...). */
export interface DrawingColor {
	kind: 'srgb' | 'scheme' | 'sys' | 'preset' | 'hsl' | 'scrgb';
	/**
	 * `srgb`: `RRGGBB`; `scheme`: scheme name (`accent1`, `tx1`, `bg1`, `phClr`, ...);
	 * `sys`: last color (`RRGGBB`) or system name; `preset`: preset name;
	 * `hsl` / `scrgb`: comma-separated components.
	 */
	value: string;
	transforms?: Array<{ name: string; val?: number }>;
}

export type ShapeFill =
	| { type: 'none' }
	| { type: 'solid'; color: DrawingColor }
	| {
			type: 'gradient';
			stops: Array<{ pos: number; color: DrawingColor }>;
			/** Linear angle in degrees (0 = left to right). */
			angle?: number;
			path?: 'circle' | 'rect' | 'shape';
	  }
	| { type: 'pattern'; preset: string; fg?: DrawingColor; bg?: DrawingColor }
	| { type: 'picture'; mediaPath: string };

export interface ShapeLine {
	fill?: ShapeFill;
	/** Width in EMU. */
	width?: number;
	/** Preset dash (`solid`, `dash`, `sysDot`, ...). */
	dash?: string;
	cap?: 'rnd' | 'sq' | 'flat';
	headEnd?: { type: string; w?: string; len?: string };
	tailEnd?: { type: string; w?: string; len?: string };
}

/** Text run properties used by shapes and charts. */
export interface DrawingTextStyle {
	/** Size in points. */
	size?: number;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean;
	strike?: boolean;
	color?: DrawingColor;
	font?: string;
}

export interface DrawingParagraph {
	align?: 'l' | 'ctr' | 'r' | 'just' | 'dist';
	runs: Array<{ text: string; style?: DrawingTextStyle }>;
}

export interface DrawingTextBody {
	paragraphs: DrawingParagraph[];
	anchor?: 't' | 'ctr' | 'b';
	/** Insets in EMU. */
	insets?: { l?: number; t?: number; r?: number; b?: number };
	wrap?: boolean;
	vertical?: string;
}

/** A cell-relative anchor point; offsets are in EMU. */
export interface AnchorPoint {
	col: number;
	colOff: number;
	row: number;
	rowOff: number;
}

export type DrawingAnchor =
	| {
			type: 'twoCell';
			from: AnchorPoint;
			to: AnchorPoint;
			editAs?: 'twoCell' | 'oneCell' | 'absolute';
	  }
	| { type: 'oneCell'; from: AnchorPoint; ext: { cx: number; cy: number } }
	| { type: 'absolute'; pos: { x: number; y: number }; ext: { cx: number; cy: number } };

export interface DrawingObjectBase {
	/** Stable id within the sheet (the `cNvPr` id as a string, or generated). */
	id: string;
	anchor: DrawingAnchor;
	name?: string;
	/** Alternative text. */
	description?: string;
	hidden?: boolean;
	/** Rotation in degrees (clockwise). */
	rotation?: number;
	flipH?: boolean;
	flipV?: boolean;
	/** Original anchor element (`xdr:twoCellAnchor`, ...) for verbatim round-trip. */
	raw?: XmlElement;
	/** The object changed since load (its anchor XML must be regenerated). */
	dirty?: boolean;
}

export interface PictureObject extends DrawingObjectBase {
	kind: 'picture';
	/** Package path of the image, a key into `workbook.media`. */
	mediaPath: string;
	/** Source crop rectangle in 1/1000ths of a percent (OOXML `srcRect`). */
	crop?: { l?: number; t?: number; r?: number; b?: number };
	line?: ShapeLine;
	/** Hyperlink target on click, when present. */
	hyperlink?: string;
}

export interface ChartObject extends DrawingObjectBase {
	kind: 'chart';
	chart: ChartModel;
	/** Package path of the chart part (`xl/charts/chart1.xml`). */
	chartPath?: string;
}

export interface ShapeObject extends DrawingObjectBase {
	kind: 'shape';
	/** Preset geometry name (`rect`, `roundRect`, `ellipse`, `rightArrow`, ...) or `custom`. */
	geometry: string;
	/** Preset geometry adjust values (`adj`, `adj1`, ...). */
	adjust?: Record<string, number>;
	fill?: ShapeFill;
	line?: ShapeLine;
	text?: DrawingTextBody;
	/** Text box shape (`txBox="1"`). */
	isTextBox?: boolean;
	/** Connector shape (`xdr:cxnSp`). */
	isConnector?: boolean;
}

export interface GroupObject extends DrawingObjectBase {
	kind: 'group';
	/**
	 * Children are positioned with `absolute` anchors expressed in the group's
	 * child coordinate space (`childOffset` / `childExtent`, EMU), which maps
	 * onto the group's own anchor box.
	 */
	children: DrawingObject[];
	childOffset?: { x: number; y: number };
	childExtent?: { cx: number; cy: number };
}

/** Content the engine cannot render (OLE frames, SmartArt, ...), preserved for round-trip. */
export interface UnknownDrawingObject extends DrawingObjectBase {
	kind: 'unknown';
	/** Short description of what the object is, for placeholders and warnings. */
	label?: string;
}

export type DrawingObject =
	| PictureObject
	| ChartObject
	| ShapeObject
	| GroupObject
	| UnknownDrawingObject;

/** A binary package part such as an image (`xl/media/image1.png`). */
export interface MediaItem {
	path: string;
	contentType: string;
	data: Uint8Array;
}
