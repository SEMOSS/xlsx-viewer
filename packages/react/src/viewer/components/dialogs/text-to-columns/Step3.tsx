import { useId } from 'react';
import type { TextColumnFormat } from 'xlsx-viewer-core';

import { useT, type TranslationKey } from '../../../../i18n';
import { FIELD, GroupBox, Radio } from '../dialog-ui';
import { DataPreview } from './DataPreview';
import { formatLabel } from './wizard-model';

const FORMATS: ReadonlyArray<{ value: TextColumnFormat; labelKey: TranslationKey }> = [
	{ value: 'general', labelKey: 'format.general' },
	{ value: 'text', labelKey: 'format.text' },
	{ value: 'skip', labelKey: 'dialog.textToColumns.skip' },
];

/** Step 3: data format of each column and the destination. */
export function Step3({
	formats,
	selected,
	onSelect,
	onFormat,
	destination,
	onDestination,
	rows,
}: {
	formats: readonly TextColumnFormat[];
	selected: number;
	onSelect: (col: number) => void;
	onFormat: (format: TextColumnFormat) => void;
	destination: string;
	onDestination: (text: string) => void;
	rows: readonly (readonly string[])[];
}) {
	const t = useT();
	const name = useId();
	const current = formats[selected] ?? 'general';
	return (
		<>
			<p>{t('dialog.textToColumns.formatIntro')}</p>
			<div className='flex flex-wrap items-start gap-4'>
				<GroupBox legend={t('dialog.textToColumns.columnFormat')} className='w-[240px]'>
					<div
						role='radiogroup'
						aria-label={t('dialog.textToColumns.columnFormat')}
						className='flex flex-col gap-1.5 pt-1'
					>
						{FORMATS.map((f) => (
							<Radio
								key={f.value}
								name={name}
								label={t(f.labelKey)}
								checked={current === f.value}
								onChange={() => onFormat(f.value)}
							/>
						))}
					</div>
				</GroupBox>
				<p className='max-w-[260px] pt-3 text-muted-foreground'>
					{t('dialog.textToColumns.generalHint')}
				</p>
			</div>
			<label className='flex items-center gap-2'>
				<span className='whitespace-nowrap'>{t('dialog.textToColumns.destination')}</span>
				<input
					className={FIELD}
					value={destination}
					onChange={(e) => onDestination(e.target.value)}
				/>
			</label>
			<DataPreview
				rows={rows}
				columns={formats.length}
				headers={formats.map((f) => formatLabel(f, t))}
				selected={selected}
				onSelect={onSelect}
			/>
		</>
	);
}
