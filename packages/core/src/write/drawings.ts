/**
 * Drawing parts (`xl/drawings/drawingN.xml`). Untouched objects are emitted
 * from their original XML (their relationship ids stay valid because the
 * original drawing rels are the starting point); new or edited objects are
 * regenerated, adding image, chart and hyperlink relationships as needed.
 */
import { writeChartXml } from '../charts';
import { imageContentType, partExtension, relativeTarget, resolveRelTarget } from '../package';
import type { ChartObject, DrawingObject } from '../types/drawing';
import type { Relationship, Worksheet } from '../types/sheet';
import { findAllDescendantsLocal, localName, parseXml, serializeElement } from '../xml';
import { CT, NS, REL } from '../xml/namespaces';
import type { XmlElement } from '../xml/types';
import { addRel, type WriteContext } from './context';
import { anchorXml, type DrawingXmlContext, objectXml } from './drawing-xml';
import { isSelfContained, prefixScope } from './raw-markup';
import { SheetGeometry } from './sheet-geometry';
import { attrString } from './xml-utils';

export interface DrawingPartPlan {
	/** Output path of the drawing part. */
	path: string;
	/** The part existed in the source (regenerated in place). */
	inPlace: boolean;
	/** The sheet is a copy: original XML and parts must not be shared. */
	fresh: boolean;
}

const DEFAULT_ROOT: Record<string, string> = { 'xmlns:xdr': NS.xdr, 'xmlns:a': NS.a };

/** Every relationship id referenced by an attribute in the markup. */
function referencedValues(xml: string): Set<string> {
	const out = new Set<string>();
	const re = /="([^"]*)"/g;
	let m: RegExpExecArray | null;
	while ((m = re.exec(xml)) !== null) {
		out.add(m[1]);
	}
	return out;
}

function collectIds(raw: XmlElement, into: Set<number>): void {
	for (const c of findAllDescendantsLocal(raw, 'cNvPr')) {
		const n = Number(c.attrs.id);
		if (Number.isInteger(n)) {
			into.add(n);
		}
	}
}

/** Charts nested anywhere in an object tree. */
function chartsOf(obj: DrawingObject): ChartObject[] {
	if (obj.kind === 'chart') {
		return [obj];
	}
	return obj.kind === 'group' ? obj.children.flatMap(chartsOf) : [];
}

function mediaTargetPath(ctx: WriteContext, mediaPath: string): string {
	if (/^xl\/media\/[^/]+\.[A-Za-z0-9]+$/.test(mediaPath) && !ctx.has(mediaPath)) {
		return mediaPath;
	}
	const item = ctx.workbook.media.get(mediaPath);
	const ext = partExtension(mediaPath) || extensionFor(item?.contentType) || 'png';
	return ctx.allocate('xl/media', 'image', ext);
}

function extensionFor(contentType: string | undefined): string | undefined {
	switch (contentType) {
		case CT.jpeg:
			return 'jpeg';
		case CT.gif:
			return 'gif';
		case CT.png:
			return 'png';
		case CT.svg:
			return 'svg';
		case CT.bmp:
			return 'bmp';
		case CT.emf:
			return 'emf';
		case CT.wmf:
			return 'wmf';
		default:
			return undefined;
	}
}

class DrawingWriter {
	readonly rels: Relationship[];
	private readonly usedIds = new Set<number>();
	private nextFree = 2;
	private readonly mediaWritten = new Map<string, string>();

	constructor(
		readonly ctx: WriteContext,
		readonly plan: DrawingPartPlan,
		readonly sheet: Worksheet,
		originalRels: Relationship[],
	) {
		this.rels = originalRels.map((r) => ({ ...r }));
	}

	reserveIds(raw: XmlElement): void {
		collectIds(raw, this.usedIds);
	}

	nextId(preferred: string): number {
		const n = Number(preferred);
		if (Number.isInteger(n) && n > 0 && !this.usedIds.has(n)) {
			this.usedIds.add(n);
			return n;
		}
		while (this.usedIds.has(this.nextFree)) {
			this.nextFree++;
		}
		this.usedIds.add(this.nextFree);
		return this.nextFree;
	}

	private relFor(type: string, path: string): string | undefined {
		const suffix = type.slice(type.lastIndexOf('/') + 1);
		return this.rels.find(
			(r) =>
				r.targetMode !== 'External' &&
				r.type.endsWith(`/${suffix}`) &&
				resolveRelTarget(this.plan.path, r.target).toLowerCase() === path.toLowerCase(),
		)?.id;
	}

	imageRel(mediaPath: string): string | undefined {
		const cached = this.mediaWritten.get(mediaPath);
		if (cached) {
			return cached;
		}
		const ctx = this.ctx;
		let path = mediaPath;
		if (!ctx.has(mediaPath)) {
			const item = ctx.workbook.media.get(mediaPath);
			if (!item) {
				return undefined;
			}
			path = mediaTargetPath(ctx, mediaPath);
			ctx.writeBytes(path, item.data);
			const ext = partExtension(path);
			ctx.writer.contentTypes.ensureDefault(ext, item.contentType || imageContentType(ext));
		}
		const id =
			this.relFor(REL.image, path) ??
			addRel(this.rels, REL.image, relativeTarget(this.plan.path, path));
		this.mediaWritten.set(mediaPath, id);
		return id;
	}

	hyperlinkRel(target: string): string {
		const internal = target.startsWith('#');
		const existing = this.rels.find(
			(r) =>
				r.type === REL.hyperlink &&
				r.target === target &&
				(r.targetMode === 'External') !== internal,
		);
		return (
			existing?.id ?? addRel(this.rels, REL.hyperlink, target, internal ? undefined : 'External')
		);
	}

	/** Write the part of a dirty chart in place (keeping the original on failure). */
	refreshChart(obj: ChartObject): void {
		if (obj.chart.dirty === false || !obj.chartPath || !this.ctx.sourceHas(obj.chartPath)) {
			return;
		}
		try {
			this.ctx.writeText(obj.chartPath, writeChartXml(obj.chart), CT.chart);
		} catch (e) {
			this.ctx.warn(
				'chart-not-written',
				`Chart ${obj.chartPath} could not be saved: ${(e as Error).message}`,
				{
					part: obj.chartPath,
					sheet: this.sheet.name,
				},
			);
		}
	}

	async chartRel(obj: ChartObject): Promise<string | undefined> {
		const ctx = this.ctx;
		const src = obj.chartPath;
		if (!this.plan.fresh && src && ctx.sourceHas(src)) {
			this.refreshChart(obj);
			return (
				this.relFor(REL.chart, src) ??
				addRel(this.rels, REL.chart, relativeTarget(this.plan.path, src))
			);
		}
		let xml: string;
		try {
			xml = writeChartXml(obj.chart);
		} catch (e) {
			ctx.warn(
				'chart-not-written',
				`Chart "${obj.name ?? obj.id}" could not be saved: ${(e as Error).message}`,
				{
					sheet: this.sheet.name,
				},
			);
			return undefined;
		}
		const path = ctx.allocate('xl/charts', 'chart', 'xml');
		ctx.writeText(path, xml, CT.chart);
		if (src && ctx.sourceHas(src) && obj.chart.dirty === false) {
			// A copied, unedited chart keeps the links of the original (style, colors, data).
			const rels = (await ctx.rels(src)).map((r) =>
				r.targetMode === 'External'
					? r
					: { ...r, target: relativeTarget(path, resolveRelTarget(src, r.target)) },
			);
			if (rels.length) {
				ctx.writeRels(path, rels);
			}
		}
		return addRel(this.rels, REL.chart, relativeTarget(this.plan.path, path));
	}
}

async function originalRoot(
	ctx: WriteContext,
	plan: DrawingPartPlan,
): Promise<Record<string, string>> {
	if (!plan.inPlace) {
		return { ...DEFAULT_ROOT };
	}
	const xml = await ctx.sourceText(plan.path);
	if (!xml) {
		return { ...DEFAULT_ROOT };
	}
	try {
		const attrs = { ...parseXml(xml).root.attrs };
		for (const [k, v] of Object.entries(DEFAULT_ROOT)) {
			if (attrs[k] === undefined) {
				attrs[k] = v;
			}
		}
		return attrs;
	} catch {
		return { ...DEFAULT_ROOT };
	}
}

/** The prefixes the generated markup relies on are bound as expected. */
function prefixesUsable(attrs: Record<string, string>): boolean {
	return attrs['xmlns:xdr'] === NS.xdr && attrs['xmlns:a'] === NS.a;
}

/**
 * Write a sheet's drawing part. Returns the number of objects written (0
 * means the caller should drop the drawing reference).
 */
export async function writeDrawingPart(
	ctx: WriteContext,
	sheet: Worksheet,
	plan: DrawingPartPlan,
): Promise<number> {
	const originalRels = plan.inPlace && !plan.fresh ? await ctx.rels(plan.path) : [];
	const w = new DrawingWriter(ctx, plan, sheet, originalRels);
	let rootAttrs = await originalRoot(ctx, plan);
	// Original markup is valid as-is in its own part; elsewhere only when self-contained.
	const trusted = plan.inPlace && !plan.fresh && prefixesUsable(rootAttrs);
	if (!prefixesUsable(rootAttrs)) {
		rootAttrs = { ...DEFAULT_ROOT };
	}
	const scope = prefixScope(rootAttrs);
	const reuse = (raw: XmlElement): boolean => trusted || isSelfContained(raw, scope);
	const verbatim = (o: DrawingObject): boolean => !!o.raw && o.dirty === false && reuse(o.raw);
	const reserve = (o: DrawingObject): void => {
		if (o.raw && (o.dirty === false || o.kind === 'unknown') && reuse(o.raw)) {
			w.reserveIds(o.raw);
		} else if (o.kind === 'group') {
			o.children.forEach(reserve);
		}
	};
	sheet.drawings.forEach(reserve);
	const geometry = new SheetGeometry(sheet);
	const pendingCharts = new Map<ChartObject, string | undefined>();
	// Chart parts are written first (async), the synchronous serializer then looks them up.
	for (const o of sheet.drawings) {
		for (const c of chartsOf(o)) {
			if (verbatim(o)) {
				w.refreshChart(c);
			} else {
				pendingCharts.set(c, await w.chartRel(c));
			}
		}
	}
	const xctx: DrawingXmlContext = {
		geometry,
		imageRel: (p) => w.imageRel(p),
		chartRel: (c) => pendingCharts.get(c),
		hyperlinkRel: (t) => w.hyperlinkRel(t),
		nextId: (p) => w.nextId(p),
		reuse,
		warn: (code, message) => ctx.warn(code, message, { sheet: sheet.name, part: plan.path }),
	};
	const parts: string[] = [];
	for (const o of sheet.drawings) {
		if (verbatim(o)) {
			parts.push(serializeElement(o.raw!));
			continue;
		}
		if (o.kind === 'unknown' && o.raw && !localName(o.raw.name).endsWith('Anchor')) {
			// Alternate content cannot be re-anchored: keep it as it was when that is valid here.
			if (reuse(o.raw)) {
				parts.push(serializeElement(o.raw));
			} else {
				xctx.warn(
					'drawing-object-dropped',
					`"${o.name ?? o.id}" (${o.label ?? 'object'}) cannot be saved here`,
				);
			}
			continue;
		}
		const inner = objectXml(o, xctx, geometry.anchorBox(o.anchor));
		if (inner) {
			parts.push(anchorXml(o.anchor, inner));
		}
	}
	if (!parts.length) {
		return 0;
	}
	const xml =
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' +
		`<xdr:wsDr${attrString(rootAttrs)}>${parts.join('')}</xdr:wsDr>`;
	ctx.writeText(plan.path, xml, CT.drawing);
	const used = referencedValues(xml);
	const rels = w.rels.filter((r) => used.has(r.id));
	await ctx.writeRelsIfChanged(plan.path, rels);
	return parts.length;
}
