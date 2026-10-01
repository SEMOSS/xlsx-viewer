import type { BorderEdge, BorderStyle, ColorContext } from 'xlsx-viewer-core';
import { resolveColorCss } from 'xlsx-viewer-core';

/** How an Excel border style is stroked (at 100% zoom). */
export interface StrokeSpec {
	/** Line width in px. */
	width: number;
	/** SVG dash pattern, when dashed. */
	dash?: string;
	/** Two parallel 1px lines with a 1px gap. */
	double?: boolean;
	/** Strength used to resolve conflicts between neighbouring cells. */
	weight: number;
}

const STROKES: Record<Exclude<BorderStyle, 'none'>, StrokeSpec> = {
	hair: { width: 1, dash: '1 1', weight: 1 },
	dotted: { width: 1, dash: '1 2', weight: 2 },
	dashDotDot: { width: 1, dash: '9 3 3 3 3 3', weight: 3 },
	dashDot: { width: 1, dash: '9 3 3 3', weight: 4 },
	dashed: { width: 1, dash: '3 1', weight: 5 },
	thin: { width: 1, weight: 6 },
	mediumDashDotDot: { width: 2, dash: '9 3 3 3 3 3', weight: 7 },
	slantDashDot: { width: 2, dash: '11 1 5 1', weight: 8 },
	mediumDashDot: { width: 2, dash: '9 3 3 3', weight: 9 },
	mediumDashed: { width: 2, dash: '9 3', weight: 10 },
	medium: { width: 2, weight: 11 },
	thick: { width: 3, weight: 12 },
	double: { width: 3, double: true, weight: 13 },
};

export function borderStroke(style: BorderStyle): StrokeSpec | undefined {
	return style === 'none' ? undefined : STROKES[style];
}

/** One drawn border line in sheet pixel coordinates. */
export interface BorderSegment {
	/** `h` = horizontal line at `y` from `x1` to `x2`; `v` = vertical line at `x` from `y1` to `y2`. */
	orientation: 'h' | 'v';
	/** Position of the line (the gridline it sits on). */
	at: number;
	from: number;
	to: number;
	color: string;
	stroke: StrokeSpec;
	zoom: number;
}

/**
 * Collects border edges for the visible cells, de-duplicating shared edges:
 * when two neighbours both draw the same gridline, the stronger style wins
 * (ties go to the right / lower cell, which Excel paints last).
 */
export class BorderCollector {
	private readonly edges = new Map<string, BorderSegment>();

	constructor(
		private readonly ctx: ColorContext,
		private readonly zoom: number,
		private readonly defaultColor = '#000000',
	) {}

	add(
		orientation: 'h' | 'v',
		/** Gridline index: row boundary for `h`, column boundary for `v`. */
		lineIndex: number,
		/** Cell index along the line: column for `h`, row for `v`. */
		cellIndex: number,
		at: number,
		from: number,
		to: number,
		edge: BorderEdge | undefined,
	): void {
		if (!edge || edge.style === 'none') {
			return;
		}
		const stroke = STROKES[edge.style];
		if (!stroke) {
			return;
		}
		const key = `${orientation}:${lineIndex}:${cellIndex}`;
		const existing = this.edges.get(key);
		if (existing && existing.stroke.weight > stroke.weight) {
			return;
		}
		this.edges.set(key, {
			orientation,
			at,
			from,
			to,
			color: resolveColorCss(edge.color, this.ctx, this.defaultColor) ?? this.defaultColor,
			stroke,
			zoom: this.zoom,
		});
	}

	segments(): BorderSegment[] {
		return [...this.edges.values()];
	}
}

/**
 * Line geometry for SVG rendering: returns one or two lines (double borders)
 * centred on the gridline. Widths scale with zoom but never drop below 1px.
 */
export function segmentLines(
	seg: BorderSegment,
): Array<{ x1: number; y1: number; x2: number; y2: number; width: number; dash?: string }> {
	const w = Math.max(1, Math.round(seg.stroke.width * Math.max(seg.zoom, 0.5)));
	const dash = seg.stroke.dash
		?.split(' ')
		.map((n) => Number(n) * Math.max(seg.zoom, 0.5))
		.join(' ');
	// Pixel-snap: 1px lines sit on the gridline pixel (x - 0.5 centre).
	const centre = seg.at - (w % 2 === 1 ? 0.5 : 0);
	if (seg.stroke.double) {
		const gap = Math.max(1, Math.round(seg.zoom));
		const offsets = [-gap, gap];
		return offsets.map((o) =>
			seg.orientation === 'h'
				? { x1: seg.from, y1: centre + o, x2: seg.to, y2: centre + o, width: 1 }
				: { x1: centre + o, y1: seg.from, x2: centre + o, y2: seg.to, width: 1 },
		);
	}
	return [
		seg.orientation === 'h'
			? { x1: seg.from, y1: centre, x2: seg.to, y2: centre, width: w, dash }
			: { x1: centre, y1: seg.from, x2: centre, y2: seg.to, width: w, dash },
	];
}
