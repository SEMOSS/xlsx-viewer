import type { PrintLayout, PrintPage, StyleResolver, Workbook } from 'xlsx-viewer-core';
import { memo, useMemo, type CSSProperties } from 'react';

import type { CellOverlays } from '../../../engine/layout/overlays';
import type { PaneLayout } from '../../../engine/layout/panes';
import { MergeIndex } from '../../../engine/layout/style-at';
import { buildPrintPageModel, type PrintHeadingCell } from '../../../engine/print/page-model';
import { createPrintGeometry } from '../../../engine/print/print-geometry';
import { fontFamilyStack } from '../../../engine/render/fonts';
import { formatNumber, useLocale, useT } from '../../../i18n';
import { DrawingObjectView } from '../drawings/DrawingObjectView';
import { GridRenderContextReact, type GridRenderContext } from '../grid/grid-context';
import { PaneLayers } from '../grid/PaneLayers';
import { HeaderFooterBand } from './HeaderFooterBand';

export interface PrintPageViewProps {
	workbook: Workbook;
	layout: PrintLayout;
	page: PrintPage;
	styles: StyleResolver;
	/** Table and conditional-formatting overlays of the sheet. */
	overlays?: CellOverlays;
	/** Preview zoom: 1 draws the page at its real size (96 DPI). */
	zoom?: number;
	/** Document revision (re-reads cells and charts after edits). */
	revision?: number;
	className?: string;
}

const NO_PANES: PaneLayout = {
	originRow: 0,
	originCol: 0,
	frozenRows: 0,
	frozenCols: 0,
	bodyRow: 0,
	bodyCol: 0,
	frozenHeight: 0,
	frozenWidth: 0,
};

/** Printed gridlines and heading borders. */
const PRINT_LINE = '#808080';

function headingStyle(cell: PrintHeadingCell, fontSize: number, fontFamily: string): CSSProperties {
	const column = cell.kind === 'col';
	const row = cell.kind === 'row';
	const line = `1px solid ${PRINT_LINE}`;
	return {
		position: 'absolute',
		left: cell.x,
		top: cell.y,
		width: cell.width,
		height: cell.height,
		boxSizing: 'border-box',
		borderRight: line,
		borderBottom: line,
		borderTop: row ? undefined : line,
		borderLeft: column ? undefined : line,
		display: 'flex',
		alignItems: 'center',
		justifyContent: 'center',
		overflow: 'hidden',
		fontSize,
		fontFamily,
		color: '#000000',
		whiteSpace: 'nowrap',
	};
}

/**
 * One printed page: white paper with the header and footer in the
 * margins, the title and body blocks (drawn by the grid's own layers),
 * gridlines and headings when they print, and the floating objects of the
 * body. Content is laid out at 96 DPI and scaled by `zoom` for previews.
 */
export const PrintPageView = memo(function PrintPageView({
	workbook,
	layout,
	page,
	styles,
	overlays,
	zoom = 1,
	revision = 0,
	className,
}: PrintPageViewProps) {
	const t = useT();
	const locale = useLocale();
	const sheetIndex = layout.sheetIndex;
	const sheet = workbook.sheets[sheetIndex];
	const geometry = useMemo(
		() => createPrintGeometry(sheet, layout.scale),
		// revision: row heights and column widths change in place.
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[sheet, layout.scale, revision],
	);
	const model = useMemo(
		() => buildPrintPageModel(layout, page, sheet, geometry),
		[layout, page, sheet, geometry],
	);
	// eslint-disable-next-line react-hooks/exhaustive-deps
	const merges = useMemo(() => new MergeIndex(sheet.merges), [sheet, revision]);
	const body = model.blocks.find((b) => b.kind === 'body');
	const ctx: GridRenderContext = useMemo(
		() => ({
			sheet,
			sheetIndex,
			geometry,
			panes: NO_PANES,
			merges,
			styles,
			colorCtx: { theme: workbook.theme, indexedColors: workbook.styles.indexedColors },
			zoom: layout.scale,
			docVersion: revision,
			date1904: workbook.date1904,
			showGridLines: model.gridLines,
			showFormulas: sheet.view.showFormulas,
			showZeros: sheet.view.showZeros,
			gridlineColor: PRINT_LINE,
			cf: overlays?.cf,
			table: overlays?.table,
			bodyRows: body?.rows ?? { start: 0, end: -1 },
			bodyCols: body?.cols ?? { start: 0, end: -1 },
		}),
		[
			sheet,
			sheetIndex,
			geometry,
			merges,
			styles,
			workbook,
			layout.scale,
			revision,
			model,
			overlays,
			body,
		],
	);
	const normal = styles.normalFont;
	const headingFont = fontFamilyStack(normal.name);
	const s = layout.scale;
	const outer: CSSProperties = {
		position: 'relative',
		width: model.width * zoom,
		height: model.height * zoom,
		overflow: 'hidden',
		background: '#ffffff',
		flexShrink: 0,
	};
	const inner: CSSProperties = {
		position: 'absolute',
		left: 0,
		top: 0,
		width: model.width,
		height: model.height,
		transform: zoom === 1 ? undefined : `scale(${zoom})`,
		transformOrigin: '0 0',
		filter: layout.blackAndWhite ? 'grayscale(1)' : undefined,
		color: '#000000',
	};
	return (
		<div
			className={`xv-print-page ${className ?? ''}`}
			style={outer}
			data-page={page.pageNumber}
			aria-label={t('print.pageNumber', { number: formatNumber(page.pageNumber, locale) })}
			role='img'
		>
			<div style={inner}>
				<GridRenderContextReact.Provider value={ctx}>
					{model.blocks.map((b) => (
						<div
							key={b.key}
							className='xv-print-block'
							data-block={b.kind}
							style={{
								position: 'absolute',
								left: b.x,
								top: b.y,
								width: b.width,
								height: b.height,
								overflow: 'hidden',
								boxShadow: model.gridLines ? `inset 1px 1px 0 0 ${PRINT_LINE}` : undefined,
							}}
						>
							<div
								style={{
									position: 'absolute',
									left: 0,
									top: 0,
									transform: `translate(${-b.originX}px, ${-b.originY}px)`,
								}}
							>
								<PaneLayers rows={b.rows} cols={b.cols} />
								{b.kind === 'body' &&
									model.drawings.map(({ object, rect }) => (
										<div
											key={object.id}
											style={{
												position: 'absolute',
												left: rect.x,
												top: rect.y,
												width: rect.width,
												height: rect.height,
											}}
										>
											<div
												style={{
													position: 'absolute',
													left: 0,
													top: 0,
													width: rect.width / s,
													height: rect.height / s,
													transform: s === 1 ? undefined : `scale(${s})`,
													transformOrigin: '0 0',
												}}
											>
												<DrawingObjectView
													object={object}
													rect={{ x: 0, y: 0, width: rect.width / s, height: rect.height / s }}
													workbook={workbook}
													sheetIndex={sheetIndex}
													zoom={1}
													revision={revision}
												/>
											</div>
										</div>
									))}
							</div>
						</div>
					))}
				</GridRenderContextReact.Provider>
				{model.headings.map((cell) => (
					<div
						key={cell.key}
						className='xv-print-heading'
						style={headingStyle(cell, model.headingFontSize, headingFont)}
					>
						{cell.text}
					</div>
				))}
				{model.header && (
					<HeaderFooterBand
						band={model.header}
						edge='top'
						pageWidth={model.width}
						scale={model.headerFooterScale}
						defaultFont={normal.name}
						colorCtx={ctx.colorCtx}
					/>
				)}
				{model.footer && (
					<HeaderFooterBand
						band={model.footer}
						edge='bottom'
						pageWidth={model.width}
						scale={model.headerFooterScale}
						defaultFont={normal.name}
						colorCtx={ctx.colorCtx}
					/>
				)}
			</div>
		</div>
	);
});
