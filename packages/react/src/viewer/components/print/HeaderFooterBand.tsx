import type { ColorContext, HeaderFooterRun } from 'xlsx-viewer-core';
import { resolveColorCss } from 'xlsx-viewer-core';
import type { CSSProperties } from 'react';

import type { PrintTextBand } from '../../../engine/print/page-model';
import { PX_PER_PT } from '../../../engine/print/print-geometry';
import { fontFamilyStack } from '../../../engine/render/fonts';

interface HeaderFooterBandProps {
	band: PrintTextBand;
	edge: 'top' | 'bottom';
	pageWidth: number;
	/** Font scale (Scale with document). */
	scale: number;
	defaultFont: string;
	colorCtx: ColorContext;
}

function runStyle(
	run: HeaderFooterRun,
	scale: number,
	defaultFont: string,
	colorCtx: ColorContext,
): CSSProperties {
	const f = run.font;
	const script = f.vertAlign !== undefined;
	const decorations = [f.underline ? 'underline' : '', f.strike ? 'line-through' : '']
		.filter(Boolean)
		.join(' ');
	return {
		fontFamily: fontFamilyStack(f.name ?? defaultFont),
		fontSize: (f.size ?? 11) * PX_PER_PT * scale * (script ? 0.7 : 1),
		fontWeight: f.bold ? 700 : 400,
		fontStyle: f.italic ? 'italic' : 'normal',
		textDecorationLine: decorations || undefined,
		textDecorationStyle: f.underline === 'double' ? 'double' : undefined,
		verticalAlign:
			f.vertAlign === 'superscript' ? 'super' : f.vertAlign === 'subscript' ? 'sub' : undefined,
		color: resolveColorCss(f.color, colorCtx) ?? '#000000',
	};
}

/** A printed header or footer: left, center and right sections between the margins. */
export function HeaderFooterBand({
	band,
	edge,
	pageWidth,
	scale,
	defaultFont,
	colorCtx,
}: HeaderFooterBandProps) {
	const base: CSSProperties = {
		position: 'absolute',
		[edge]: band.offset,
		whiteSpace: 'pre',
		lineHeight: 1.2,
		fontSize: 11 * PX_PER_PT * scale,
		fontFamily: fontFamilyStack(defaultFont),
		color: '#000000',
	};
	const sections: Array<[keyof PrintTextBand['sections'], CSSProperties]> = [
		['left', { left: band.left, textAlign: 'left' }],
		['center', { left: 0, width: pageWidth, textAlign: 'center' }],
		['right', { right: pageWidth - band.right, textAlign: 'right' }],
	];
	return (
		<div className={`xv-print-${edge === 'top' ? 'header' : 'footer'}`}>
			{sections.map(([name, pos]) => {
				const runs = band.sections[name];
				if (runs.length === 0) {
					return null;
				}
				return (
					<div key={name} data-section={name} style={{ ...base, ...pos }}>
						{runs.map((run, i) => (
							// biome-ignore lint: runs are positional
							<span key={i} style={runStyle(run, scale, defaultFont, colorCtx)}>
								{run.text}
							</span>
						))}
					</div>
				);
			})}
		</div>
	);
}
