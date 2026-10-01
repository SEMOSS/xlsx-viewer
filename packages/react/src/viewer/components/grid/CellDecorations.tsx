import type { DataValidation, FilterTarget, Hyperlink, Note } from 'xlsx-viewer-core';
import { findValidation, getListItems, hyperlinkDescription, setCellInput } from 'xlsx-viewer-core';
import { ChevronDown, Funnel } from 'lucide-react';
import { memo, useRef, useState } from 'react';

import { useT } from '../../../i18n';
import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import type { Span } from '../../../engine/layout/types';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { Popover } from '../ui/Popover';
import { FilterMenu } from './FilterMenu';

interface Props {
	geometry: SheetGeometry;
	rows: Span;
	cols: Span;
	hoverNote: Note | null;
	/** Hyperlink cell under the pointer (tooltip). */
	hoverLink?: HoverLink | null;
}

export interface HoverLink {
	link: Hyperlink;
	row: number;
	col: number;
}

interface FilterButton {
	key: string;
	row: number;
	col: number;
	target: FilterTarget;
	colId: number;
	active: boolean;
}

function inSpan(row: number, col: number, rows: Span, cols: Span): boolean {
	return row >= rows.start && row <= rows.end && col >= cols.start && col <= cols.end;
}

/** Note markers and popups, AutoFilter buttons and the validation dropdown arrow, for one pane. */
export const CellDecorations = memo(function CellDecorations({
	geometry: g,
	rows,
	cols,
	hoverNote,
	hoverLink,
}: Props) {
	const t = useT();
	const store = useViewerStore();
	const snap = useStoreSelector((s) => s);
	const sheet = store.workbook.sheets[snap.ui.activeSheet];
	const sel = store.selectionOf(snap.ui.activeSheet);
	const [openFilter, setOpenFilter] = useState<FilterButton | null>(null);
	const [listOpen, setListOpen] = useState(false);
	const filterAnchor = useRef<HTMLButtonElement | null>(null);
	const listAnchor = useRef<HTMLButtonElement | null>(null);

	const buttons: FilterButton[] = [];
	if (sheet.autoFilter) {
		const r = sheet.autoFilter.ref;
		for (let c = r.left; c <= r.right; c++) {
			buttons.push({
				key: `af:${c}`,
				row: r.top,
				col: c,
				target: { kind: 'sheet' },
				colId: c - r.left,
				active: sheet.autoFilter.columns.some((x) => x.colId === c - r.left),
			});
		}
	}
	sheet.tables.forEach((table, index) => {
		if (!table.headerRowCount || table.autoFilter === undefined) {
			return;
		}
		for (let c = table.ref.left; c <= table.ref.right; c++) {
			buttons.push({
				key: `t${index}:${c}`,
				row: table.ref.top,
				col: c,
				target: { kind: 'table', index },
				colId: c - table.ref.left,
				active: table.autoFilter.columns.some((x) => x.colId === c - table.ref.left),
			});
		}
	});

	const notes = sheet.notes.filter((n) => inSpan(n.ref.row, n.ref.col, rows, cols));
	const comments = sheet.threadedComments.filter(
		(c) => !c.parentId && inSpan(c.ref.row, c.ref.col, rows, cols),
	);
	const size = Math.max(5, Math.round(6 * g.zoom));

	// Validation list dropdown on the active cell.
	const { row: ar, col: ac } = sel.active;
	const dv: DataValidation | undefined =
		!store.readOnly && inSpan(ar, ac, rows, cols) ? findValidation(sheet, ar, ac) : undefined;
	const showList = dv?.type === 'list' && dv.inCellDropdown !== false;
	// Data validation input message ("Input Message" tab), shown while the cell is active.
	const inputDv = inSpan(ar, ac, rows, cols) ? findValidation(sheet, ar, ac) : undefined;
	const inputMessage =
		inputDv?.showInputMessage && (inputDv.promptTitle || inputDv.prompt) ? inputDv : undefined;
	const listItems = () =>
		dv
			? getListItems(dv, {
					evaluate: (f, s, r, c) => store.calc.evaluateFormula(f, s, r, c),
					sheetIndex: snap.ui.activeSheet,
					row: ar,
					col: ac,
					date1904: store.workbook.date1904,
					workbook: store.workbook,
				} as Parameters<typeof getListItems>[1])
			: [];

	return (
		<div className='xv-decorations' style={{ position: 'absolute', left: 0, top: 0, zIndex: 13 }}>
			{notes.map((n) => {
				const right = g.colLeft(n.ref.col) + g.colWidth(n.ref.col);
				return (
					<span
						key={`n:${n.ref.row}:${n.ref.col}`}
						className='xv-note-marker'
						style={{
							left: right - size - 1,
							top: g.rowTop(n.ref.row),
							borderWidth: `0 ${size}px ${size}px 0`,
						}}
					/>
				);
			})}
			{comments.map((c) => {
				const right = g.colLeft(c.ref.col) + g.colWidth(c.ref.col);
				return (
					<span
						key={`c:${c.id}`}
						className='xv-comment-marker'
						style={{
							left: right - size - 2,
							top: g.rowTop(c.ref.row) + 1,
							width: size + 1,
							height: size + 1,
						}}
					/>
				);
			})}
			{[
				...notes.filter((n) => n.visible),
				...(hoverNote && !hoverNote.visible ? [hoverNote] : []),
			].map((n) => (
				<div
					key={`np:${n.ref.row}:${n.ref.col}`}
					className='xv-note-popup'
					style={{
						left: g.colLeft(n.ref.col) + g.colWidth(n.ref.col) + 12,
						top: Math.max(0, g.rowTop(n.ref.row) - 4),
						width: n.width ?? 180,
						minHeight: n.height ? Math.min(n.height, 260) : 60,
						fontSize: `${Math.max(9, 11 * g.zoom)}px`,
					}}
				>
					{n.author && <strong>{n.author}:</strong>} {n.text}
				</div>
			))}
			{hoverLink && !hoverNote && inSpan(hoverLink.row, hoverLink.col, rows, cols) && (
				<div
					className='xv-link-tip'
					role='tooltip'
					style={{
						left: g.colLeft(hoverLink.col) + Math.min(24, g.colWidth(hoverLink.col) / 2),
						top: g.rowTop(hoverLink.row) + g.rowHeight(hoverLink.row) + 6,
					}}
				>
					<div className='xv-link-tip-target'>{hyperlinkDescription(hoverLink.link)}</div>
					<div>{t('hyperlink.followHint')}</div>
				</div>
			)}
			{buttons
				.filter((b) => inSpan(b.row, b.col, rows, cols))
				.map((b) => {
					const right = g.colLeft(b.col) + g.colWidth(b.col);
					const h = Math.min(g.rowHeight(b.row) - 2, Math.round(17 * g.zoom));
					return (
						<button
							key={b.key}
							type='button'
							data-xv-object='filter-button'
							aria-label={t('filter.columnButton', { column: b.colId + 1 })}
							className='xv-filter-button'
							data-active={b.active || undefined}
							style={{
								left: right - h - 2,
								top: g.rowTop(b.row) + g.rowHeight(b.row) - h - 2,
								width: h,
								height: h,
							}}
							onPointerDown={(e) => {
								e.stopPropagation();
								filterAnchor.current = e.currentTarget;
								setOpenFilter((cur) => (cur?.key === b.key ? null : b));
							}}
						>
							{b.active ? (
								<Funnel size={Math.max(9, h - 7)} />
							) : (
								<ChevronDown size={Math.max(9, h - 5)} />
							)}
						</button>
					);
				})}
			{inputMessage && !snap.ui.edit && (
				<div
					className='xv-input-message'
					role='tooltip'
					style={{
						left: g.colLeft(ac) + Math.min(g.colWidth(ac), 40),
						top: g.rowTop(ar) + g.rowHeight(ar) + 4,
						fontSize: `${Math.max(10, 11 * g.zoom)}px`,
					}}
				>
					{inputMessage.promptTitle && (
						<strong className='block'>{inputMessage.promptTitle}</strong>
					)}
					{inputMessage.prompt}
				</div>
			)}
			{showList && (
				<button
					type='button'
					data-xv-object='validation-list'
					aria-label={t('grid.chooseFromList')}
					className='xv-list-button'
					style={{
						left: g.colLeft(ac) + g.colWidth(ac) + 1,
						top: g.rowTop(ar) + g.rowHeight(ar) - Math.round(18 * g.zoom),
						width: Math.round(16 * g.zoom),
						height: Math.round(18 * g.zoom),
					}}
					onPointerDown={(e) => {
						e.stopPropagation();
						listAnchor.current = e.currentTarget;
						setListOpen((v) => !v);
					}}
				>
					<ChevronDown size={12} />
				</button>
			)}
			<Popover
				open={openFilter !== null}
				onClose={() => setOpenFilter(null)}
				anchorRef={filterAnchor}
				ariaLabel={t('cmd.filter')}
			>
				{openFilter && (
					<FilterMenu
						target={openFilter.target}
						colId={openFilter.colId}
						onClose={() => setOpenFilter(null)}
					/>
				)}
			</Popover>
			<Popover
				open={listOpen && showList}
				onClose={() => setListOpen(false)}
				anchorRef={listAnchor}
				role='listbox'
				ariaLabel={t('grid.validationList')}
			>
				<div className='max-h-[240px] min-w-[140px] overflow-y-auto py-1 text-[13px]'>
					{listOpen &&
						listItems().map((item) => (
							<button
								key={item}
								type='button'
								role='option'
								className='block w-full px-3 py-1 text-left hover:bg-accent'
								onClick={() => {
									setListOpen(false);
									store.transact(store.t('undo.typing', { text: item }), (tx) => {
										setCellInput(tx, store.workbook, snap.ui.activeSheet, ar, ac, item);
									});
								}}
							>
								{item}
							</button>
						))}
				</div>
			</Popover>
		</div>
	);
});
