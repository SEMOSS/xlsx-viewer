import { a1, type IconSetName } from 'xlsx-viewer-core';
import { memo, type CSSProperties, type ReactNode } from 'react';

import { CELL_PADDING_X } from '../../../engine/layout/viewport-layout';
import type { TextBox } from '../../../engine/layout/types';
import type { CellContent } from '../../../engine/render/cell-content';
import { IconSetIcon } from '../cf/IconSetIcon';

const LINE_HEIGHT = 1.25;

function fontStyle(c: CellContent, scale: number): CSSProperties {
	return {
		fontFamily: c.font.fontFamily,
		fontSize: `${c.font.fontSize * scale}px`,
		fontWeight: c.font.fontWeight,
		fontStyle: c.font.fontStyle,
		color: c.color,
		textDecorationLine: c.font.textDecorationLine,
		textDecorationStyle: c.font.textDecorationStyle,
		verticalAlign: c.font.verticalAlign,
		lineHeight: LINE_HEIGHT,
	};
}

function renderRuns(c: CellContent, scale: number): ReactNode {
	if (!c.runs) {
		return c.text;
	}
	return c.runs.map((run, i) => (
		<span
			// biome-ignore lint: runs are positional and never reordered
			key={i}
			style={{
				fontFamily: run.font.fontFamily,
				fontSize: `${run.font.fontSize * scale}px`,
				fontWeight: run.font.fontWeight,
				fontStyle: run.font.fontStyle,
				color: run.color,
				textDecorationLine: run.font.textDecorationLine,
				textDecorationStyle: run.font.textDecorationStyle,
				verticalAlign: run.font.verticalAlign,
			}}
		>
			{run.text}
		</span>
	));
}

const V_JUSTIFY: Record<CellContent['vAlign'], CSSProperties['alignItems']> = {
	top: 'flex-start',
	center: 'center',
	bottom: 'flex-end',
	justify: 'flex-start',
	distributed: 'center',
};

/**
 * One cell's text. Single-line text is absolutely positioned inside its clip
 * box so it can overflow into empty neighbours with Excel's alignment;
 * wrapped, rotated and filled text use a flex box sized to the cell.
 */
export const CellTextView = memo(function CellTextView({
	box,
	zoom,
}: {
	box: TextBox;
	zoom: number;
}) {
	const c = box.content;
	const padX = Math.max(1, CELL_PADDING_X * zoom);
	const { clip, cell } = box;
	// `data-cell` names the cell a text box draws (a test and tooling contract).
	const ref = a1(box.row, box.col);
	const outer: CSSProperties = {
		position: 'absolute',
		left: clip.x,
		top: clip.y,
		width: clip.width,
		height: clip.height,
		overflow: 'hidden',
		pointerEvents: 'none',
	};
	const icon = box.icon ? (
		<span className='xv-cf-icon' style={{ left: Math.max(1, 2 * zoom) }}>
			<IconSetIcon
				iconSet={box.icon.iconSet as IconSetName}
				index={box.icon.index}
				size={Math.round(16 * zoom)}
			/>
		</span>
	) : null;

	if (c.repeatFill && !box.hashes) {
		// Accounting formats: "$" at the left edge, the number at the right.
		return (
			<div
				data-cell={ref}
				style={{
					...outer,
					display: 'flex',
					alignItems: V_JUSTIFY[c.vAlign],
					padding: `0 ${padX}px`,
				}}
			>
				<div style={{ ...fontStyle(c, 1), display: 'flex', width: '100%', whiteSpace: 'pre' }}>
					<span>{c.repeatFill.before}</span>
					<span style={{ flex: 1, overflow: 'hidden' }}>
						{c.repeatFill.char === ' ' ? '' : c.repeatFill.char.repeat(200)}
					</span>
					<span>{c.repeatFill.after}</span>
				</div>
			</div>
		);
	}

	const complex =
		c.wrap || c.vertical || c.rotation !== 0 || c.hAlign === 'fill' || c.hAlign === 'distributed';
	if (complex) {
		const textAlign: CSSProperties['textAlign'] =
			c.hAlign === 'right'
				? 'right'
				: c.hAlign === 'center' || c.hAlign === 'centerContinuous'
					? 'center'
					: c.hAlign === 'justify' || c.hAlign === 'distributed'
						? 'justify'
						: 'left';
		const inner: CSSProperties = {
			...fontStyle(c, box.shrinkScale),
			whiteSpace: c.wrap ? 'pre-wrap' : 'pre',
			overflowWrap: 'anywhere',
			textAlign,
			width: c.rotation !== 0 ? undefined : '100%',
			paddingLeft: padX + (c.hAlign === 'left' ? c.indent : 0),
			paddingRight: padX + (c.hAlign === 'right' ? c.indent : 0),
			writingMode: c.vertical ? 'vertical-rl' : undefined,
			textOrientation: c.vertical ? 'upright' : undefined,
			transform: c.rotation !== 0 ? `rotate(${c.rotation}deg)` : undefined,
		};
		const text =
			c.hAlign === 'fill' && c.text
				? c.text.repeat(
						Math.max(
							1,
							Math.floor(cell.width / Math.max(1, c.text.length * c.font.fontSize * 0.55)),
						),
					)
				: null;
		return (
			<div
				data-cell={ref}
				style={{
					...outer,
					display: 'flex',
					alignItems: V_JUSTIFY[c.vAlign],
					justifyContent:
						c.rotation !== 0 || c.vertical
							? textAlign === 'right'
								? 'flex-end'
								: textAlign === 'center'
									? 'center'
									: 'flex-start'
							: undefined,
				}}
			>
				{icon}
				<div style={inner}>{text ?? renderRuns(c, box.shrinkScale)}</div>
			</div>
		);
	}

	// Single line: position relative to the cell inside the (possibly wider) clip.
	const span: CSSProperties = {
		...fontStyle(c, box.shrinkScale),
		position: 'absolute',
		whiteSpace: 'pre',
	};
	const cellLeft = cell.x - clip.x;
	const cellRight = clip.x + clip.width - (cell.x + cell.width);
	if (c.hAlign === 'right') {
		span.right = cellRight + padX + c.indent;
	} else if (c.hAlign === 'center' || c.hAlign === 'centerContinuous') {
		span.left = cellLeft + cell.width / 2;
		span.transform = 'translateX(-50%)';
	} else {
		span.left = cellLeft + padX + c.indent + (box.icon ? 18 * zoom : 0);
	}
	if (c.vAlign === 'top') {
		span.top = 1;
	} else if (c.vAlign === 'center' || c.vAlign === 'distributed') {
		span.top = '50%';
		span.transform = `${span.transform ?? ''} translateY(-50%)`.trim();
	} else {
		span.bottom = Math.max(0, Math.round(1 * zoom));
	}
	return (
		<div data-cell={ref} style={outer}>
			{icon}
			<span style={span}>{renderRuns(c, box.shrinkScale)}</span>
		</div>
	);
});
