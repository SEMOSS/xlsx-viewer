import { useT } from '../../../../i18n';
import { toggleSheetViewFlag } from '../../../commands/view-commands';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import { printFlagCmd } from '../../print/page-setup-commands';
import { RibbonGroup } from '../RibbonGroup';

interface OptionRow {
	label: string;
	view: 'showGridLines' | 'showRowColHeaders';
	print: 'gridLines' | 'headings';
}

/** Page Layout > Sheet Options: show and print gridlines and headings. */
export function SheetOptionsGroup() {
	const t = useT();
	const store = useViewerStore();
	const sheet = useStoreSelector(
		(s, st) => (void s.docVersion, st.workbook.sheets[s.ui.activeSheet]),
		() => false,
	);
	const rows: OptionRow[] = [
		{ label: t('cmd.gridlines'), view: 'showGridLines', print: 'gridLines' },
		{ label: t('cmd.headings'), view: 'showRowColHeaders', print: 'headings' },
	];
	const box = 'h-3 w-3 accent-[var(--xlsx-selection-border,#107c41)]';
	return (
		<RibbonGroup label={t('group.sheetOptions')}>
			<div className='grid grid-cols-[auto_auto_auto] items-center gap-x-2 gap-y-0.5 text-[11.5px]'>
				{rows.map((r) => (
					<div key={r.print} className='contents'>
						<span className='font-semibold text-foreground'>{r.label}</span>
						<label className='flex items-center gap-1 whitespace-nowrap'>
							<input
								type='checkbox'
								className={box}
								aria-label={t('pageLayout.viewOption', { option: r.label })}
								checked={sheet.view[r.view]}
								onChange={() => toggleSheetViewFlag(store, r.view)}
							/>
							{t('cmd.viewOption')}
						</label>
						<label className='flex items-center gap-1 whitespace-nowrap'>
							<input
								type='checkbox'
								className={box}
								aria-label={t('pageLayout.printOption', { option: r.label })}
								checked={sheet.printOptions?.[r.print] === true}
								disabled={store.readOnly}
								onChange={(e) => printFlagCmd(store, r.print, e.target.checked)}
							/>
							{t('cmd.printOption')}
						</label>
					</div>
				))}
			</div>
		</RibbonGroup>
	);
}
