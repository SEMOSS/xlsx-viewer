/** Maps one chart render-tree node onto SVG elements. */
import type { SVGProps } from 'react';
import type {
	ChartNodeMeta,
	ChartPaint,
	ChartRenderNode,
	ChartTextNode,
	ChartTextRun,
} from 'xlsx-viewer-core';

/** Ids assigned to the gradients and clip paths of a tree (see `ChartSvg`). */
export interface ChartSvgIds {
	gradient: WeakMap<object, string>;
	clip: WeakMap<object, string>;
}

const round = (v: number): number => Math.round(v * 100) / 100;

/**
 * Snap the cross coordinate of a horizontal / vertical hairline to the pixel
 * grid (half pixels for odd widths) so gridlines and axes render crisp.
 */
function snap(v: number, width: number): number {
	const w = Math.max(1, Math.round(width));
	return w % 2 === 1 ? Math.floor(v) + 0.5 : Math.round(v);
}

function lineCoords(node: Extract<ChartRenderNode, { type: 'line' }>) {
	const w = node.strokeWidth ?? 1;
	if (w <= 2 && node.x1 === node.x2) {
		const x = snap(node.x1, w);
		return { x1: x, y1: round(node.y1), x2: x, y2: round(node.y2) };
	}
	if (w <= 2 && node.y1 === node.y2) {
		const y = snap(node.y1, w);
		return { x1: round(node.x1), y1: y, x2: round(node.x2), y2: y };
	}
	return { x1: round(node.x1), y1: round(node.y1), x2: round(node.x2), y2: round(node.y2) };
}

function paintProps(node: ChartPaint, ids: ChartSvgIds): SVGProps<SVGElement> {
	const gradientId = node.gradient ? ids.gradient.get(node.gradient) : undefined;
	const props: SVGProps<SVGElement> = {
		fill: gradientId ? `url(#${gradientId})` : (node.fill ?? 'none'),
	};
	if (node.stroke) {
		props.stroke = node.stroke;
		props.strokeWidth = round(node.strokeWidth ?? 1);
		if (node.dash?.length) {
			props.strokeDasharray = node.dash.map(round).join(' ');
		}
		if (node.lineCap) {
			props.strokeLinecap = node.lineCap;
		}
		if (node.lineJoin) {
			props.strokeLinejoin = node.lineJoin;
		}
	}
	if (node.opacity !== undefined && node.opacity < 1) {
		props.opacity = node.opacity;
	}
	return props;
}

function metaProps(node: ChartNodeMeta): Record<string, string | number | undefined> {
	return {
		'data-role': node.role,
		'data-plot': node.plot,
		'data-series': node.series,
		'data-point': node.point,
	};
}

function TextNode({ node }: { node: ChartTextNode }) {
	const transform =
		node.rotation !== undefined && node.rotation !== 0
			? `rotate(${round(node.rotation)} ${round(node.pivotX ?? node.x)} ${round(node.pivotY ?? node.y)})`
			: undefined;
	const lines: ChartTextRun[][] = node.runs ?? node.lines.map((text) => [{ text }]);
	return (
		<text
			x={round(node.x)}
			y={round(node.y)}
			fontSize={round(node.fontSize)}
			fontFamily={node.fontFamily}
			fontWeight={node.fontWeight}
			fontStyle={node.fontStyle}
			textDecoration={node.decoration}
			fill={node.fill}
			textAnchor={node.anchor}
			transform={transform}
			style={{ whiteSpace: 'pre' }}
			{...metaProps(node)}
		>
			{lines.map((runs, i) =>
				runs.map((run, j) => (
					<tspan
						key={`${i}-${j}`}
						x={j === 0 ? round(node.x) : undefined}
						dy={j === 0 && i > 0 ? round(node.lineHeight) : undefined}
						fontSize={run.fontSize !== undefined ? round(run.fontSize) : undefined}
						fontWeight={run.fontWeight}
						fontStyle={run.fontStyle}
						textDecoration={run.decoration}
						fill={run.fill}
						fontFamily={run.fontFamily}
					>
						{run.text}
					</tspan>
				)),
			)}
		</text>
	);
}

/** Render a node (recursively for groups). */
export function ChartSvgNode({ node, ids }: { node: ChartRenderNode; ids: ChartSvgIds }) {
	switch (node.type) {
		case 'group': {
			const clipId = node.clip ? ids.clip.get(node.clip) : undefined;
			return (
				<g clipPath={clipId ? `url(#${clipId})` : undefined} {...metaProps(node)}>
					{node.children.map((child, i) => (
						<ChartSvgNode key={i} node={child} ids={ids} />
					))}
				</g>
			);
		}
		case 'rect':
			return (
				<rect
					x={round(node.x)}
					y={round(node.y)}
					width={round(Math.max(0, node.width))}
					height={round(Math.max(0, node.height))}
					rx={node.rx ? round(node.rx) : undefined}
					{...(paintProps(node, ids) as SVGProps<SVGRectElement>)}
					{...metaProps(node)}
				/>
			);
		case 'path':
			return (
				<path
					d={node.d}
					{...(paintProps(node, ids) as SVGProps<SVGPathElement>)}
					{...metaProps(node)}
				/>
			);
		case 'line':
			return (
				<line
					{...lineCoords(node)}
					{...(paintProps(node, ids) as SVGProps<SVGLineElement>)}
					{...metaProps(node)}
				/>
			);
		case 'polyline': {
			const pts: string[] = [];
			for (let i = 0; i + 1 < node.points.length; i += 2) {
				pts.push(`${round(node.points[i])},${round(node.points[i + 1])}`);
			}
			return (
				<polyline
					points={pts.join(' ')}
					{...(paintProps(node, ids) as SVGProps<SVGPolylineElement>)}
					{...metaProps(node)}
				/>
			);
		}
		case 'circle':
			return (
				<circle
					cx={round(node.cx)}
					cy={round(node.cy)}
					r={round(Math.max(0, node.r))}
					{...(paintProps(node, ids) as SVGProps<SVGCircleElement>)}
					{...metaProps(node)}
				/>
			);
		case 'text':
			return <TextNode node={node} />;
		default:
			return null;
	}
}
