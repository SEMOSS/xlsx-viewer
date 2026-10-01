import type {
	ConditionalFormatEngine,
	PrintLayout,
	RangeAddress,
	StyleResolver,
	Workbook,
} from 'xlsx-viewer-core';
import { paginateSheets } from 'xlsx-viewer-core';
import { useMemo } from 'react';
import { flushSync } from 'react-dom';
import { createRoot } from 'react-dom/client';

import { createCellOverlays } from '../../../engine/layout/overlays';
import { ViewerStoreContext } from '../../state/use-store';
import { ViewerStore } from '../../state/viewer-store';
import { PrintPageView } from './PrintPageView';

export interface PrintDocumentProps {
	workbook: Workbook;
	layouts: PrintLayout[];
	styles: StyleResolver;
	conditionalFormats?: ConditionalFormatEngine | null;
	revision?: number;
}

/** CSS page name for a paper size (pages of several sizes use named `@page` rules). */
export function pageName(layout: PrintLayout): string {
	return `xvp${Math.round(layout.paper.width)}x${Math.round(layout.paper.height)}`;
}

/** Every page of a print job at real size, one element per printed page. */
export function PrintDocument({
	workbook,
	layouts,
	styles,
	conditionalFormats,
	revision = 0,
}: PrintDocumentProps) {
	const overlays = useMemo(
		() =>
			layouts.map((l) => createCellOverlays(workbook, l.sheetIndex, conditionalFormats ?? null)),
		[workbook, layouts, conditionalFormats],
	);
	return (
		<div className='xv-print-document'>
			{layouts.map((layout, i) =>
				layout.pages.map((page) => (
					<div
						key={`${layout.sheetIndex}:${page.index}`}
						className={`xv-print-sheet ${pageName(layout)}`}
						style={{
							width: `${layout.paper.width}pt`,
							// A hair under the paper height keeps browsers from adding blank pages.
							height: `calc(${layout.paper.height}pt - 1px)`,
							overflow: 'hidden',
							position: 'relative',
						}}
					>
						<PrintPageView
							workbook={workbook}
							layout={layout}
							page={page}
							styles={styles}
							overlays={overlays[i]}
							revision={revision}
						/>
					</div>
				)),
			)}
		</div>
	);
}

/** `@page` rules and print CSS for a job. */
export function printCss(layouts: readonly PrintLayout[]): string {
	const first = layouts.find((l) => l.pageCount > 0) ?? layouts[0];
	const size = (l: PrintLayout) => `${l.paper.width}pt ${l.paper.height}pt`;
	const rules = [
		first ? `@page { size: ${size(first)}; margin: 0; }` : '@page { margin: 0; }',
		'html, body { margin: 0; padding: 0; background: #ffffff; }',
		'* { -webkit-print-color-adjust: exact; print-color-adjust: exact; }',
		'.xv-print-sheet { break-after: page; page-break-after: always; break-inside: avoid; }',
		'.xv-print-sheet:last-child { break-after: auto; page-break-after: auto; }',
	];
	const seen = new Set<string>();
	for (const l of layouts) {
		const name = pageName(l);
		if (!seen.has(name)) {
			seen.add(name);
			rules.push(`@page ${name} { size: ${size(l)}; margin: 0; }`, `.${name} { page: ${name}; }`);
		}
	}
	return rules.join('\n');
}

export interface PrintWorkbookOptions {
	/** Sheets to print, in order (default: the active sheet). */
	sheets?: number[];
	/** Print Selection: print this range of the (single) sheet instead of its print area. */
	range?: RangeAddress;
	/** Several ranges (a multi-area selection), each on its own pages. */
	ranges?: RangeAddress[];
	ignorePrintArea?: boolean;
	/** For `&F` in headers and footers. */
	fileName?: string;
	/** The viewer's store (styles, conditional formats); a read-only one is created otherwise. */
	store?: ViewerStore;
	revision?: number;
	/** Opens the print dialog (default: focus the frame and `print()`); tests replace it. */
	print?: (win: Window) => void;
}

function whenLoaded(el: HTMLElement): Promise<void> {
	return new Promise<void>((resolve) => {
		el.addEventListener('load', () => resolve(), { once: true });
		el.addEventListener('error', () => resolve(), { once: true });
	});
}

/** Images and copied stylesheets of the print frame have loaded (or `timeoutMs` passed). */
function waitForResources(doc: Document, timeoutMs: number): Promise<void> {
	const pending: Array<Promise<void>> = [];
	for (const img of Array.from(doc.images)) {
		if (!img.complete) {
			pending.push(whenLoaded(img));
		}
	}
	for (const link of Array.from(doc.querySelectorAll<HTMLLinkElement>('link[rel="stylesheet"]'))) {
		if (!link.sheet) {
			pending.push(whenLoaded(link));
		}
	}
	if (pending.length === 0) {
		return Promise.resolve();
	}
	return Promise.race([
		Promise.all(pending).then(() => undefined),
		new Promise<void>((resolve) => setTimeout(resolve, timeoutMs)),
	]);
}

function copyStyles(from: Document, to: Document): void {
	for (const node of Array.from(from.querySelectorAll('style, link[rel="stylesheet"]'))) {
		to.head.appendChild(to.importNode(node, true));
	}
}

/**
 * Print sheets with Excel's page layout: every page is rendered into a
 * hidden iframe (`@page` sized to the paper, no browser margins, one
 * element per page) and the browser's print dialog opens on it (which also
 * offers Save as PDF). Returns the number of pages printed (0 when there is
 * nothing to print).
 */
export async function printWorkbook(
	workbook: Workbook,
	options: PrintWorkbookOptions = {},
): Promise<number> {
	const sheets = options.sheets?.length ? options.sheets : [workbook.view.activeTab];
	const layouts = paginateSheets(workbook, sheets, {
		ranges: options.ranges ?? (options.range ? [options.range] : undefined),
		ignorePrintArea: options.ignorePrintArea,
		fileName: options.fileName,
	});
	const pages = layouts.reduce((n, l) => n + l.pageCount, 0);
	if (pages === 0 || typeof document === 'undefined') {
		return 0;
	}
	const store = options.store ?? new ViewerStore(workbook, { readOnly: true });
	const iframe = document.createElement('iframe');
	iframe.setAttribute('aria-hidden', 'true');
	iframe.title = store.t('print.title');
	iframe.style.cssText = 'position:fixed;right:0;bottom:0;width:0;height:0;border:0';
	document.body.appendChild(iframe);
	const doc = iframe.contentDocument;
	const win = iframe.contentWindow;
	if (!doc || !win) {
		iframe.remove();
		return 0;
	}
	doc.open();
	doc.write('<!doctype html><html><head><meta charset="utf-8"></head><body></body></html>');
	doc.close();
	doc.title = layouts[0]?.sheetName ?? store.t('print.documentTitle');
	copyStyles(document, doc);
	const style = doc.createElement('style');
	style.textContent = printCss(layouts);
	doc.head.appendChild(style);
	const container = doc.createElement('div');
	doc.body.appendChild(container);
	const root = createRoot(container);
	flushSync(() =>
		root.render(
			// Grid layers read the workbook from the store (sparklines).
			<ViewerStoreContext.Provider value={store}>
				<PrintDocument
					workbook={workbook}
					layouts={layouts}
					styles={store.styles}
					conditionalFormats={store.conditionalFormats}
					revision={options.revision}
				/>
			</ViewerStoreContext.Provider>,
		),
	);
	await waitForResources(doc, 3000);
	await (doc as Document & { fonts?: { ready?: Promise<unknown> } }).fonts?.ready;
	let cleaned = false;
	const cleanup = () => {
		if (!cleaned) {
			cleaned = true;
			root.unmount();
			iframe.remove();
		}
	};
	win.addEventListener('afterprint', () => setTimeout(cleanup, 0), { once: true });
	const openDialog =
		options.print ??
		((w: Window) => {
			w.focus();
			w.print();
		});
	openDialog(win);
	// Browsers without `afterprint` (or a blocked dialog): clean up later.
	setTimeout(cleanup, 60_000);
	return pages;
}
