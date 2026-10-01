import { Minus, Plus, Table2 } from 'lucide-react';
import { useEffect, useMemo } from 'react';

import { formatNumber, formatPercent, useLocale, useT } from '../../../i18n';
import { formatStatistic } from '../../../engine/selection/stat-format';
import { selectionStats } from '../../../engine/selection/stats';
import { styleIdAt } from '../../../engine/layout/style-at';
import { canPoint } from '../../commands/point-commands';
import { setZoom } from '../../commands/view-commands';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { useReturnToGrid, useShell } from '../../shell-context';

const ZOOM_STEPS = [10, 25, 50, 75, 90, 100, 110, 125, 150, 175, 200, 250, 300, 400];

/** Bottom bar: mode, messages, selection statistics and zoom. */
export function StatusBar() {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const shell = useShell();
	const returnToGrid = useReturnToGrid();
	const snap = useStoreSelector((s) => s);
	const sheet = store.workbook.sheets[snap.ui.activeSheet];
	const sel = store.selectionOf(snap.ui.activeSheet);
	const zoom = sheet.view.zoomScale || 100;
	const edit = snap.ui.edit;
	const mode = edit
		? canPoint(edit) && edit.point
			? t('status.point')
			: edit.mode === 'enter'
				? t('status.enter')
				: t('status.edit')
		: t('status.ready');

	const stats = useMemo(
		() => {
			const single =
				sel.ranges.length === 1 &&
				sel.ranges[0].top === sel.ranges[0].bottom &&
				sel.ranges[0].left === sel.ranges[0].right;
			return single ? null : selectionStats(sheet, sel.ranges);
		},
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[sheet, sel, snap.docVersion],
	);

	useEffect(() => {
		if (!snap.ui.message) {
			return;
		}
		const timer = setTimeout(() => store.setUi({ message: null }), 4000);
		return () => clearTimeout(timer);
	}, [snap.ui.message, store]);

	// Statistics use the active cell's number format, like Excel; General follows the UI locale.
	const fmt = store.styles.resolve(styleIdAt(sheet, sel.active.row, sel.active.col)).numFmt;
	const show = (n: number | null) =>
		n === null ? '' : formatStatistic(n, fmt, locale, store.workbook.date1904);

	return (
		<footer className='xv-status flex h-6 shrink-0 items-center gap-3 border-t border-border bg-card px-3 text-[11.5px] text-muted-foreground'>
			<span className='min-w-[44px] text-foreground'>{mode}</span>
			{snap.ui.message && (
				<span role='status' className='truncate rounded-sm bg-accent px-2 text-foreground'>
					{snap.ui.message}
				</span>
			)}
			<div className='flex-1' />
			{stats && stats.count > 0 && (
				<div className='flex items-center gap-3 tabular-nums'>
					{stats.numericCount > 0 && (
						<span>
							{t('status.average')}
							{t('ui.labelSeparator')}
							<span className='text-foreground'>{show(stats.average)}</span>
						</span>
					)}
					<span>
						{t('status.count')}
						{t('ui.labelSeparator')}
						<span className='text-foreground'>{formatNumber(stats.count, locale)}</span>
					</span>
					{stats.numericCount > 0 && (
						<span>
							{t('status.sum')}
							{t('ui.labelSeparator')}
							<span className='text-foreground'>{show(stats.sum)}</span>
						</span>
					)}
				</div>
			)}
			{snap.ui.manualCalc && (
				<button
					type='button'
					className='rounded-sm px-1.5 text-foreground hover:bg-accent'
					onClick={() => store.recalculate()}
				>
					{t('status.calculate')}
				</button>
			)}
			<span title={t('status.normalView')} className='text-foreground'>
				<Table2 size={13} />
			</span>
			<div className='flex items-center gap-1'>
				<button
					type='button'
					aria-label={t('status.zoomOut')}
					className='rounded-sm p-0.5 hover:bg-accent'
					onClick={() => setZoom(store, [...ZOOM_STEPS].reverse().find((z) => z < zoom) ?? 10)}
				>
					<Minus size={12} />
				</button>
				<input
					type='range'
					min={10}
					max={400}
					step={5}
					value={zoom}
					aria-label={t('cmd.zoom')}
					className='xv-zoom-slider h-1 w-[100px] cursor-pointer accent-[var(--xlsx-primary)]'
					onChange={(e) => setZoom(store, Number(e.target.value))}
					// Dragging the slider leaves the keyboard with the grid, as in Excel.
					onPointerUp={(e) => returnToGrid(e.currentTarget)}
				/>
				<button
					type='button'
					aria-label={t('status.zoomIn')}
					className='rounded-sm p-0.5 hover:bg-accent'
					onClick={() => setZoom(store, ZOOM_STEPS.find((z) => z > zoom) ?? 400)}
				>
					<Plus size={12} />
				</button>
				<button
					type='button'
					className='min-w-[42px] whitespace-nowrap rounded-sm text-right tabular-nums hover:bg-accent'
					onClick={() => shell.openDialog('zoom')}
				>
					{formatPercent(zoom, locale)}
				</button>
			</div>
		</footer>
	);
}
