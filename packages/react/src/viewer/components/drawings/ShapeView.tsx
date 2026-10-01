import type { DrawingParagraph, ShapeFill, ShapeLine, ShapeObject, Theme } from 'xlsx-viewer-core';
import { EMU_PER_PIXEL, EMU_PER_POINT, resolveDrawingColorCss } from 'xlsx-viewer-core';
import type { CSSProperties } from 'react';

import { isLinePreset, presetPath } from '../../../engine/render/shape-paths';

function fillCss(fill: ShapeFill | undefined, theme: Theme): string {
	if (!fill) {
		return 'none';
	}
	switch (fill.type) {
		case 'none':
			return 'none';
		case 'solid':
			return resolveDrawingColorCss(fill.color, theme) ?? 'none';
		case 'gradient':
			return resolveDrawingColorCss(fill.stops[0]?.color, theme) ?? 'none';
		case 'pattern':
			return resolveDrawingColorCss(fill.fg, theme) ?? 'none';
		default:
			return 'none';
	}
}

function lineCss(
	line: ShapeLine | undefined,
	theme: Theme,
	zoom: number,
): { stroke: string; width: number; dash?: string } {
	if (!line || line.fill?.type === 'none') {
		return { stroke: 'none', width: 0 };
	}
	const stroke = line.fill ? fillCss(line.fill, theme) : '#000000';
	const width = Math.max(0.75, ((line.width ?? 9525) / EMU_PER_PIXEL) * zoom);
	const dash =
		line.dash && line.dash !== 'solid'
			? line.dash.includes('dot') || line.dash === 'sysDot'
				? `${width} ${width * 2}`
				: `${width * 4} ${width * 3}`
			: undefined;
	return { stroke, width, dash };
}

function paragraphStyle(p: DrawingParagraph): CSSProperties {
	return {
		textAlign:
			p.align === 'ctr'
				? 'center'
				: p.align === 'r'
					? 'right'
					: p.align === 'just'
						? 'justify'
						: 'left',
		margin: 0,
	};
}

/** A preset-geometry shape or text box with its text. */
export function ShapeView({
	shape,
	width,
	height,
	theme,
	zoom,
}: {
	shape: ShapeObject;
	width: number;
	height: number;
	theme: Theme;
	zoom: number;
}) {
	const line = lineCss(shape.line, theme, zoom);
	const isLine = isLinePreset(shape.geometry);
	const fill = isLine ? 'none' : fillCss(shape.fill, theme);
	const d = presetPath(shape.geometry, Math.max(1, width), Math.max(1, height), shape.adjust);
	const body = shape.text;
	const insets = body?.insets ?? {};
	const pad = (v: number | undefined, def: number) => ((v ?? def) / EMU_PER_PIXEL) * zoom;
	return (
		<>
			<svg
				width={width}
				height={height}
				style={{ position: 'absolute', inset: 0, overflow: 'visible' }}
				aria-hidden
			>
				<path
					d={d}
					fill={fill}
					stroke={line.stroke}
					strokeWidth={line.width}
					strokeDasharray={line.dash}
				/>
			</svg>
			{body && body.paragraphs.some((p) => p.runs.some((r) => r.text)) && (
				<div
					style={{
						position: 'absolute',
						inset: 0,
						display: 'flex',
						flexDirection: 'column',
						justifyContent:
							body.anchor === 'ctr' ? 'center' : body.anchor === 'b' ? 'flex-end' : 'flex-start',
						padding: `${pad(insets.t, 45720)}px ${pad(insets.r, 91440)}px ${pad(insets.b, 45720)}px ${pad(insets.l, 91440)}px`,
						overflow: 'hidden',
						whiteSpace: body.wrap === false ? 'pre' : 'pre-wrap',
						fontFamily: 'Calibri, Carlito, sans-serif',
						lineHeight: 1.2,
						pointerEvents: 'none',
					}}
				>
					{body.paragraphs.map((p, i) => (
						// biome-ignore lint: paragraphs are positional
						<p key={i} style={paragraphStyle(p)}>
							{p.runs.length === 0 ? ' ' : null}
							{p.runs.map((r, j) => (
								<span
									// biome-ignore lint: runs are positional
									key={j}
									style={{
										fontSize: `${((r.style?.size ?? 11) * EMU_PER_POINT * zoom) / EMU_PER_PIXEL}px`,
										fontWeight: r.style?.bold ? 700 : 400,
										fontStyle: r.style?.italic ? 'italic' : 'normal',
										textDecoration:
											[r.style?.underline ? 'underline' : '', r.style?.strike ? 'line-through' : '']
												.join(' ')
												.trim() || undefined,
										color:
											resolveDrawingColorCss(r.style?.color, theme) ??
											(shape.isTextBox ? '#000000' : '#ffffff'),
										fontFamily: r.style?.font
											? `'${r.style.font}', Calibri, sans-serif`
											: undefined,
									}}
								>
									{r.text}
								</span>
							))}
						</p>
					))}
				</div>
			)}
		</>
	);
}
