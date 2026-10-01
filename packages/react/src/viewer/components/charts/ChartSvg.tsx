/** Renders a chart render tree (from the core layout engine) as inline SVG. */
import { useId, useMemo } from 'react';
import type { CSSProperties } from 'react';
import type { ChartGradient, ChartRenderNode, ChartRenderTree } from 'xlsx-viewer-core';
import { ChartSvgNode, type ChartSvgIds } from './ChartSvgNode';

export interface ChartSvgProps {
	tree: ChartRenderTree;
	className?: string;
	style?: CSSProperties;
}

interface Defs {
	ids: ChartSvgIds;
	gradients: Array<{ id: string; gradient: ChartGradient }>;
	clips: Array<{ id: string; box: { x: number; y: number; width: number; height: number } }>;
}

function collectDefs(root: ChartRenderNode, prefix: string): Defs {
	const defs: Defs = {
		ids: { gradient: new WeakMap(), clip: new WeakMap() },
		gradients: [],
		clips: [],
	};
	const visit = (node: ChartRenderNode): void => {
		if (node.type === 'group') {
			if (node.clip) {
				const id = `${prefix}-clip-${defs.clips.length}`;
				defs.ids.clip.set(node.clip, id);
				defs.clips.push({ id, box: node.clip });
			}
			node.children.forEach(visit);
			return;
		}
		if (node.type !== 'text' && node.gradient && !defs.ids.gradient.has(node.gradient)) {
			const id = `${prefix}-grad-${defs.gradients.length}`;
			defs.ids.gradient.set(node.gradient, id);
			defs.gradients.push({ id, gradient: node.gradient });
		}
	};
	visit(root);
	return defs;
}

function GradientDef({ id, gradient }: { id: string; gradient: ChartGradient }) {
	const stops = gradient.stops.map((s, i) => (
		<stop
			key={i}
			offset={`${Math.round(Math.max(0, Math.min(1, s.offset)) * 10000) / 100}%`}
			stopColor={s.color}
		/>
	));
	if (gradient.kind === 'radial') {
		return (
			<radialGradient id={id} cx='50%' cy='50%' r='50%'>
				{stops}
			</radialGradient>
		);
	}
	const rad = (gradient.angle * Math.PI) / 180;
	const dx = Math.cos(rad) * 50;
	const dy = Math.sin(rad) * 50;
	const pct = (v: number) => `${Math.round(v * 100) / 100}%`;
	return (
		<linearGradient id={id} x1={pct(50 - dx)} y1={pct(50 - dy)} x2={pct(50 + dx)} y2={pct(50 + dy)}>
			{stops}
		</linearGradient>
	);
}

/** Inline SVG for a laid-out chart. */
export function ChartSvg({ tree, className, style }: ChartSvgProps) {
	const prefix = `xlsx-chart-${useId().replace(/[^a-zA-Z0-9_-]/g, '')}`;
	const defs = useMemo(() => collectDefs(tree.root, prefix), [tree, prefix]);
	return (
		<svg
			xmlns='http://www.w3.org/2000/svg'
			role='img'
			aria-label={tree.description}
			className={className}
			style={style}
			width={tree.width}
			height={tree.height}
			viewBox={`0 0 ${tree.width} ${tree.height}`}
			shapeRendering='geometricPrecision'
			textRendering='geometricPrecision'
		>
			{defs.gradients.length || defs.clips.length ? (
				<defs>
					{defs.gradients.map((g) => (
						<GradientDef key={g.id} id={g.id} gradient={g.gradient} />
					))}
					{defs.clips.map((c) => (
						<clipPath key={c.id} id={c.id}>
							<rect
								x={c.box.x}
								y={c.box.y}
								width={Math.max(0, c.box.width)}
								height={Math.max(0, c.box.height)}
							/>
						</clipPath>
					))}
				</defs>
			) : null}
			<ChartSvgNode node={tree.root} ids={defs.ids} />
		</svg>
	);
}
