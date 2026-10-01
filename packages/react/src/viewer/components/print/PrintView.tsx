import type { PrintLayout, PrintPage } from 'xlsx-viewer-core';
import { paginateSheets } from 'xlsx-viewer-core';
import { ChevronLeft, ChevronRight, Maximize2 } from 'lucide-react';
import { useEffect, useMemo, useRef, useState } from 'react';

import { formatNumber, useLocale, useT } from '../../../i18n';
import { createCellOverlays } from '../../../engine/layout/overlays';
import { PX_PER_PT } from '../../../engine/print/print-geometry';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { printWorkbook } from './print-document';
import { PrintPageView } from './PrintPageView';
import { PrintSettings, type PrintWhat } from './PrintSettings';

interface PreviewEntry {
	layout: PrintLayout;
	page: PrintPage;
}

/** Size of the preview area (ResizeObserver; a fixed guess where unavailable). */
function useAreaSize(ref: React.RefObject<HTMLDivElement | null>): {
	width: number;
	height: number;
} {
	const [size, setSize] = useState({ width: 640, height: 720 });
	useEffect(() => {
		const el = ref.current;
		if (!el || typeof ResizeObserver === 'undefined') {
			return;
		}
		const update = () => {
			if (el.clientWidth > 0 && el.clientHeight > 0) {
				setSize({ width: el.clientWidth, height: el.clientHeight });
			}
		};
		update();
		const ro = new ResizeObserver(update);
		ro.observe(el);
		return () => ro.disconnect();
	}, [ref]);
	return size;
}

/**
 * File > Print: settings on the left (applied to the active sheet as
 * undoable edits, like Excel), the print preview on the right with a page
 * navigator and Zoom to Page.
 */
export function PrintView({ fileName }: { fileName: string }) {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const docVersion = useStoreSelector((s) => s.docVersion);
	const sheetIndex = useStoreSelector((s) => s.ui.activeSheet);
	const selection = useStoreSelector((s, st) => st.selectionOf(s.ui.activeSheet));
	const [what, setWhat] = useState<PrintWhat>('active');
	const [ignorePrintArea, setIgnorePrintArea] = useState(false);
	const [copies, setCopies] = useState(1);
	const [pageIndex, setPageIndex] = useState(0);
	const [zoomed, setZoomed] = useState(false);
	const areaRef = useRef<HTMLDivElement>(null);
	const area = useAreaSize(areaRef);
	const workbook = store.workbook;
	const sheet = workbook.sheets[sheetIndex];

	const sheets = useMemo(
		() =>
			what === 'workbook'
				? workbook.sheets.flatMap((s, i) =>
						s.state === 'visible' && s.kind === 'worksheet' ? [i] : [],
					)
				: [sheetIndex],
		[what, workbook, sheetIndex],
	);
	const ranges = what === 'selection' ? selection.ranges : undefined;
	const layouts = useMemo(
		() => paginateSheets(workbook, sheets, { ranges, ignorePrintArea, fileName }),
		// docVersion: page setup and cells change in place.
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[workbook, sheets, ranges, ignorePrintArea, fileName, docVersion],
	);
	const entries = useMemo<PreviewEntry[]>(
		() => layouts.flatMap((layout) => layout.pages.map((page) => ({ layout, page }))),
		[layouts],
	);
	const count = entries.length;
	const current = entries[Math.min(pageIndex, Math.max(0, count - 1))];
	useEffect(() => {
		if (pageIndex > 0 && pageIndex >= count) {
			setPageIndex(Math.max(0, count - 1));
		}
	}, [count, pageIndex]);

	const overlaySheet = current?.layout.sheetIndex ?? sheetIndex;
	const overlays = useMemo(
		() => createCellOverlays(workbook, overlaySheet, store.conditionalFormats),
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[workbook, overlaySheet, store, docVersion],
	);

	const pageW = current ? current.layout.paper.width * PX_PER_PT : 816;
	const pageH = current ? current.layout.paper.height * PX_PER_PT : 1056;
	const fit = Math.max(0.1, Math.min((area.width - 48) / pageW, (area.height - 48) / pageH));
	const zoom = zoomed ? 1 : Math.min(1, fit);

	const print = () => {
		void printWorkbook(workbook, {
			sheets,
			ranges,
			ignorePrintArea,
			fileName,
			store,
			revision: docVersion,
		});
	};
	const go = (index: number) => setPageIndex(Math.max(0, Math.min(count - 1, index)));

	return (
		<div className='xv-print-view flex min-h-0 min-w-0 flex-1 gap-6 p-8'>
			<PrintSettings
				store={store}
				sheet={sheet}
				what={what}
				onWhat={(w) => {
					setWhat(w);
					setPageIndex(0);
				}}
				ignorePrintArea={ignorePrintArea}
				onIgnorePrintArea={setIgnorePrintArea}
				copies={copies}
				onCopies={setCopies}
				pageCount={count}
				onPrint={print}
			/>
			<div className='flex min-h-0 min-w-0 flex-1 flex-col gap-2'>
				<div
					ref={areaRef}
					className='xv-print-preview flex min-h-0 flex-1 overflow-auto rounded-md bg-muted p-6'
					aria-label={t('print.preview')}
				>
					{current ? (
						<div className='m-auto shadow-[0_1px_6px_rgba(0,0,0,0.25)]'>
							<PrintPageView
								workbook={workbook}
								layout={current.layout}
								page={current.page}
								styles={store.styles}
								overlays={overlays}
								zoom={zoom}
								revision={docVersion}
							/>
						</div>
					) : (
						<p className='m-auto text-[13px] text-muted-foreground'>{t('print.nothing')}</p>
					)}
				</div>
				<div className='flex items-center justify-between text-[13px]'>
					<nav aria-label={t('print.pages')} className='flex items-center gap-1.5'>
						<button
							type='button'
							aria-label={t('print.previousPage')}
							disabled={pageIndex <= 0}
							onClick={() => go(pageIndex - 1)}
							className='rounded-sm p-1 hover:bg-accent disabled:opacity-40'
						>
							<ChevronLeft size={16} />
						</button>
						<input
							aria-label={t('print.currentPage')}
							className='h-6 w-10 rounded-sm border border-input bg-card text-center text-[13px]'
							value={count ? Math.min(pageIndex, count - 1) + 1 : 0}
							onChange={(e) => {
								const n = Number(e.target.value);
								if (Number.isFinite(n) && n >= 1) {
									go(n - 1);
								}
							}}
						/>
						<span>{t('print.ofPages', { count: formatNumber(count, locale) })}</span>
						<button
							type='button'
							aria-label={t('print.nextPage')}
							disabled={pageIndex >= count - 1}
							onClick={() => go(pageIndex + 1)}
							className='rounded-sm p-1 hover:bg-accent disabled:opacity-40'
						>
							<ChevronRight size={16} />
						</button>
					</nav>
					<button
						type='button'
						aria-pressed={zoomed}
						title={t('print.zoomToPage')}
						onClick={() => setZoomed((z) => !z)}
						className={`flex items-center gap-1.5 rounded-sm px-2 py-1 hover:bg-accent ${zoomed ? 'bg-accent text-primary' : ''}`}
					>
						<Maximize2 size={14} /> {t('print.zoomToPage')}
					</button>
				</div>
			</div>
		</div>
	);
}
