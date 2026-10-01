import { paginateSheet } from 'xlsx-viewer-core';
import { useEffect, useMemo, useState, type KeyboardEvent } from 'react';

import { formatNumber, useLocale, useT } from '../../../../i18n';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import { fitToCmd, scaleCmd } from '../../print/page-setup-commands';
import { RibbonGroup } from '../RibbonGroup';

const SMALL =
	'h-[20px] rounded-sm border border-input bg-card px-1 text-[11.5px] text-foreground outline-none focus:border-primary disabled:opacity-50';

const PAGE_CHOICES = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];

/** Page Layout > Scale to Fit: pages wide / tall (Automatic or 1-9) and the scale. */
export function ScaleToFitGroup() {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const docVersion = useStoreSelector((s) => s.docVersion);
	const sheetIndex = useStoreSelector((s) => s.ui.activeSheet);
	const sheet = store.workbook.sheets[sheetIndex];
	const setup = sheet.pageSetup;
	const fitting = setup?.fitToPage === true;
	const width = fitting ? (setup?.fitToWidth ?? 1) : 0;
	const height = fitting ? (setup?.fitToHeight ?? 1) : 0;
	const automatic = width === 0 && height === 0;
	// While fitting, the box shows the scale Excel computes (read-only).
	const fittedScale = useMemo(
		() =>
			fitting && !automatic
				? Math.round(paginateSheet(store.workbook, sheetIndex).scale * 100)
				: null,
		// eslint-disable-next-line react-hooks/exhaustive-deps
		[store, sheetIndex, fitting, automatic, docVersion],
	);
	const scale = fittedScale ?? Math.round(setup?.scale ?? 100);
	const [text, setText] = useState(String(scale));
	useEffect(() => setText(String(scale)), [scale]);
	const ro = store.readOnly;
	const label = (n: number) =>
		n === 0
			? t('pageLayout.automatic')
			: n === 1
				? t('pageLayout.onePage')
				: t('pageLayout.pages', { count: formatNumber(n, locale) });
	const commit = () => {
		const n = Number(text.replace('%', '').trim());
		if (Number.isFinite(n) && n > 0 && Math.round(n) !== scale) {
			scaleCmd(store, n);
		} else {
			setText(String(scale));
		}
	};
	const onKey = (e: KeyboardEvent<HTMLInputElement>) => {
		if (e.key === 'Enter') {
			e.preventDefault();
			commit();
		} else if (e.key === 'Escape') {
			setText(String(scale));
		}
	};
	const choices = (current: number) =>
		PAGE_CHOICES.includes(current) ? PAGE_CHOICES : [...PAGE_CHOICES, current];
	return (
		<RibbonGroup label={t('group.scaleToFit')}>
			<div className='grid grid-flow-col grid-rows-2 items-center gap-x-2 gap-y-0.5 text-[11.5px]'>
				<label className='flex items-center justify-between gap-1.5'>
					{t('cmd.fitWidth')}
					<select
						aria-label={t('cmd.fitWidth')}
						className={SMALL}
						value={width}
						disabled={ro}
						onChange={(e) => fitToCmd(store, 'width', Number(e.target.value))}
					>
						{choices(width).map((n) => (
							<option key={n} value={n}>
								{label(n)}
							</option>
						))}
					</select>
				</label>
				<label className='flex items-center justify-between gap-1.5'>
					{t('cmd.fitHeight')}
					<select
						aria-label={t('cmd.fitHeight')}
						className={SMALL}
						value={height}
						disabled={ro}
						onChange={(e) => fitToCmd(store, 'height', Number(e.target.value))}
					>
						{choices(height).map((n) => (
							<option key={n} value={n}>
								{label(n)}
							</option>
						))}
					</select>
				</label>
				<label className='flex items-center gap-1.5'>
					{t('cmd.scale')}
					<span className='flex items-center gap-0.5'>
						<input
							aria-label={t('cmd.scale')}
							className={`${SMALL} w-[44px] text-right`}
							value={text}
							disabled={ro || !automatic}
							onChange={(e) => setText(e.target.value)}
							onBlur={commit}
							onKeyDown={onKey}
						/>
						%
					</span>
				</label>
			</div>
		</RibbonGroup>
	);
}
