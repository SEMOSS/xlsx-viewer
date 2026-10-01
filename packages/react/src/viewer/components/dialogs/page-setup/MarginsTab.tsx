import type { MarginKey } from '../../../../engine/print/page-setup-draft';
import { useT, type TranslationKey } from '../../../../i18n';
import { Checkbox, FIELD_SM, GroupBox } from '../dialog-ui';
import type { DraftTabProps } from './PageTab';

const FIELDS: ReadonlyArray<{ key: MarginKey; labelKey: TranslationKey }> = [
	{ key: 'top', labelKey: 'dialog.pageSetup.top' },
	{ key: 'header', labelKey: 'dialog.pageSetup.header' },
	{ key: 'left', labelKey: 'dialog.pageSetup.left' },
	{ key: 'right', labelKey: 'dialog.pageSetup.right' },
	{ key: 'bottom', labelKey: 'dialog.pageSetup.bottom' },
	{ key: 'footer', labelKey: 'dialog.pageSetup.footer' },
];

/** Page Setup > Margins: the six margins in inches and centering. */
export function MarginsTab({ draft, update }: DraftTabProps) {
	const t = useT();
	return (
		<div className='flex flex-col gap-3'>
			<div className='grid grid-cols-2 gap-x-6 gap-y-2'>
				{FIELDS.map(({ key, labelKey }) => (
					<label key={key} className='flex items-center justify-between gap-2'>
						<span>{t(labelKey)}</span>
						<span className='flex items-center gap-1'>
							<input
								aria-label={t(labelKey)}
								className={`${FIELD_SM} w-20 text-right`}
								inputMode='decimal'
								value={draft.margins[key]}
								onChange={(e) => update({ margins: { ...draft.margins, [key]: e.target.value } })}
							/>
							<span className='text-muted-foreground'>{t('dialog.pageSetup.inches')}</span>
						</span>
					</label>
				))}
			</div>
			<GroupBox legend={t('dialog.pageSetup.centerOnPage')}>
				<div className='flex gap-6 pt-1'>
					<Checkbox
						label={t('dialog.pageSetup.horizontally')}
						checked={draft.centerH}
						onChange={(v) => update({ centerH: v })}
					/>
					<Checkbox
						label={t('dialog.pageSetup.vertically')}
						checked={draft.centerV}
						onChange={(v) => update({ centerV: v })}
					/>
				</div>
			</GroupBox>
		</div>
	);
}
