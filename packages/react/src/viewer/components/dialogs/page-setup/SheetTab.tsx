import { useId } from 'react';

import { useT } from '../../../../i18n';
import { Checkbox, FIELD, GroupBox, Radio } from '../dialog-ui';
import type { DraftTabProps } from './PageTab';

/** Page Setup > Sheet: print area, print titles, print options and page order. */
export function SheetTab({ draft, update }: DraftTabProps) {
	const t = useT();
	const name = useId();
	return (
		<div className='flex flex-col gap-3'>
			<label className='flex items-center gap-2'>
				<span className='w-[170px] shrink-0'>{t('dialog.pageSetup.printArea')}</span>
				<input
					className={FIELD}
					value={draft.printArea}
					placeholder='$A$1:$D$20'
					onChange={(e) => update({ printArea: e.target.value })}
				/>
			</label>
			<GroupBox legend={t('dialog.pageSetup.printTitles')}>
				<div className='flex flex-col gap-2 pt-1'>
					<label className='flex items-center gap-2'>
						<span className='w-[160px] shrink-0'>{t('dialog.pageSetup.rowsAtTop')}</span>
						<input
							className={FIELD}
							value={draft.titleRows}
							placeholder='$1:$1'
							onChange={(e) => update({ titleRows: e.target.value })}
						/>
					</label>
					<label className='flex items-center gap-2'>
						<span className='w-[160px] shrink-0'>{t('dialog.pageSetup.colsAtLeft')}</span>
						<input
							className={FIELD}
							value={draft.titleCols}
							placeholder='$A:$A'
							onChange={(e) => update({ titleCols: e.target.value })}
						/>
					</label>
				</div>
			</GroupBox>
			<GroupBox legend={t('dialog.pageSetup.print')}>
				<div className='grid grid-cols-2 gap-1.5 pt-1'>
					<Checkbox
						label={t('dialog.pageSetup.gridlines')}
						checked={draft.gridLines}
						onChange={(v) => update({ gridLines: v })}
					/>
					<Checkbox
						label={t('dialog.pageSetup.headings')}
						checked={draft.headings}
						onChange={(v) => update({ headings: v })}
					/>
					<Checkbox
						label={t('dialog.pageSetup.blackAndWhite')}
						checked={draft.blackAndWhite}
						onChange={(v) => update({ blackAndWhite: v })}
					/>
					<Checkbox
						label={t('dialog.pageSetup.draftQuality')}
						checked={draft.draft}
						onChange={(v) => update({ draft: v })}
					/>
				</div>
			</GroupBox>
			<GroupBox legend={t('dialog.pageSetup.pageOrder')}>
				<div className='flex gap-6 pt-1'>
					<Radio
						name={`${name}-order`}
						label={t('dialog.pageSetup.downThenOver')}
						checked={draft.pageOrder === 'downThenOver'}
						onChange={() => update({ pageOrder: 'downThenOver' })}
					/>
					<Radio
						name={`${name}-order`}
						label={t('dialog.pageSetup.overThenDown')}
						checked={draft.pageOrder === 'overThenDown'}
						onChange={() => update({ pageOrder: 'overThenDown' })}
					/>
				</div>
			</GroupBox>
		</div>
	);
}
