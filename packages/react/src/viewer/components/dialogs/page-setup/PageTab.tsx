import { PAPER_SIZES, paperSizeDetail } from 'xlsx-viewer-core';
import { useId } from 'react';

import type { PageSetupDraft } from '../../../../engine/print/page-setup-draft';
import { useT } from '../../../../i18n';
import { FIELD, FIELD_SM, GroupBox, Radio } from '../dialog-ui';

export interface DraftTabProps {
	draft: PageSetupDraft;
	update: (patch: Partial<PageSetupDraft>) => void;
}

/** Page Setup > Page: orientation, scaling, paper size and first page number. */
export function PageTab({ draft, update }: DraftTabProps) {
	const t = useT();
	const name = useId();
	const fit = draft.scaleMode === 'fit';
	return (
		<div className='flex flex-col gap-3'>
			<GroupBox legend={t('dialog.pageSetup.orientation')}>
				<div className='flex gap-6 pt-1'>
					<Radio
						name={`${name}-o`}
						label={t('dialog.pageSetup.portrait')}
						checked={draft.orientation === 'portrait'}
						onChange={() => update({ orientation: 'portrait' })}
					/>
					<Radio
						name={`${name}-o`}
						label={t('dialog.pageSetup.landscape')}
						checked={draft.orientation === 'landscape'}
						onChange={() => update({ orientation: 'landscape' })}
					/>
				</div>
			</GroupBox>
			<GroupBox legend={t('dialog.pageSetup.scaling')}>
				<div className='flex flex-col gap-2 pt-1'>
					<div className='flex items-center gap-2'>
						<Radio
							name={`${name}-s`}
							label={t('dialog.pageSetup.adjustTo')}
							checked={!fit}
							onChange={() => update({ scaleMode: 'adjust' })}
						/>
						<input
							aria-label={t('dialog.pageSetup.adjustToPercent')}
							className={`${FIELD_SM} w-16`}
							value={draft.scale}
							disabled={fit}
							onChange={(e) => update({ scale: e.target.value })}
						/>
						<span>{t('dialog.pageSetup.percentNormalSize')}</span>
					</div>
					<div className='flex flex-wrap items-center gap-2'>
						<Radio
							name={`${name}-s`}
							label={t('dialog.pageSetup.fitTo')}
							checked={fit}
							onChange={() => update({ scaleMode: 'fit' })}
						/>
						<input
							aria-label={t('dialog.pageSetup.pagesWide')}
							className={`${FIELD_SM} w-14`}
							value={draft.fitWidth}
							disabled={!fit}
							onChange={(e) => update({ fitWidth: e.target.value })}
						/>
						<span>{t('dialog.pageSetup.pagesWideBy')}</span>
						<input
							aria-label={t('dialog.pageSetup.pagesTall')}
							className={`${FIELD_SM} w-14`}
							value={draft.fitHeight}
							disabled={!fit}
							onChange={(e) => update({ fitHeight: e.target.value })}
						/>
						<span>{t('dialog.pageSetup.tall')}</span>
					</div>
					<p className='text-[12px] text-muted-foreground'>
						{t('dialog.pageSetup.emptyIsAutomatic')}
					</p>
				</div>
			</GroupBox>
			<label className='flex items-center gap-2'>
				<span className='w-[130px] shrink-0'>{t('dialog.pageSetup.paperSize')}</span>
				<select
					className={FIELD}
					value={draft.paperSize}
					onChange={(e) => update({ paperSize: Number(e.target.value) })}
				>
					{PAPER_SIZES.map((p) => (
						<option key={p.code} value={p.code}>
							{`${p.name} (${paperSizeDetail(p)})`}
						</option>
					))}
				</select>
			</label>
			<label className='flex items-center gap-2'>
				<span className='w-[130px] shrink-0'>{t('dialog.pageSetup.firstPageNumber')}</span>
				<input
					className={`${FIELD_SM} w-24`}
					value={draft.firstPageNumber}
					onChange={(e) => update({ firstPageNumber: e.target.value })}
				/>
			</label>
		</div>
	);
}
