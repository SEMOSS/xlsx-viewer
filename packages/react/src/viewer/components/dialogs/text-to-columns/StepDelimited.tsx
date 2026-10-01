import { useT, type TranslationKey } from '../../../../i18n';
import { Checkbox, FIELD_SM, GroupBox, Select, type SelectOption } from '../dialog-ui';
import { DataPreview } from './DataPreview';
import type { DelimiterState } from './wizard-model';

const BOXES: ReadonlyArray<{
	key: 'tab' | 'semicolon' | 'comma' | 'space';
	labelKey: TranslationKey;
}> = [
	{ key: 'tab', labelKey: 'dialog.textToColumns.tab' },
	{ key: 'semicolon', labelKey: 'dialog.textToColumns.semicolon' },
	{ key: 'comma', labelKey: 'dialog.textToColumns.comma' },
	{ key: 'space', labelKey: 'dialog.textToColumns.space' },
];

/** Step 2 for delimited data: delimiters, consecutive delimiters and the text qualifier. */
export function StepDelimited({
	value,
	onChange,
	rows,
	columns,
}: {
	value: DelimiterState;
	onChange: (patch: Partial<DelimiterState>) => void;
	rows: readonly (readonly string[])[];
	columns: number;
}) {
	const t = useT();
	const qualifiers: readonly SelectOption<DelimiterState['qualifier']>[] = [
		{ value: '"', label: '"' },
		{ value: "'", label: "'" },
		// Excel writes "no qualifier" in braces: {none}.
		{ value: 'none', label: `{${t('dialog.textToColumns.noQualifier')}}` },
	];
	return (
		<>
			<p>{t('dialog.textToColumns.delimitersIntro')}</p>
			<div className='flex flex-wrap items-start gap-4'>
				<GroupBox legend={t('dialog.textToColumns.delimiters')} className='w-[170px]'>
					<div className='flex flex-col gap-1.5 pt-1'>
						{BOXES.map((b) => (
							<Checkbox
								key={b.key}
								label={t(b.labelKey)}
								checked={value[b.key]}
								onChange={(on) => onChange({ [b.key]: on })}
							/>
						))}
						<div className='flex items-center gap-1.5'>
							<Checkbox
								label={t('dialog.textToColumns.other')}
								checked={value.other}
								onChange={(on) => onChange({ other: on })}
							/>
							<input
								aria-label={t('dialog.textToColumns.otherDelimiter')}
								className={`${FIELD_SM} w-9 text-center`}
								maxLength={1}
								value={value.otherChar}
								onChange={(e) => {
									const otherChar = e.target.value.slice(-1);
									onChange({ otherChar, other: otherChar !== '' });
								}}
							/>
						</div>
					</div>
				</GroupBox>
				<div className='flex flex-col gap-3 pt-3'>
					<Checkbox
						label={t('dialog.textToColumns.consecutive')}
						checked={value.consecutive}
						onChange={(on) => onChange({ consecutive: on })}
					/>
					<label className='flex items-center gap-2'>
						<span className='whitespace-nowrap'>{t('dialog.textToColumns.qualifierLabel')}</span>
						<Select
							label={t('dialog.textToColumns.qualifier')}
							className='w-24'
							value={value.qualifier}
							options={qualifiers}
							onChange={(qualifier) => onChange({ qualifier })}
						/>
					</label>
				</div>
			</div>
			<DataPreview rows={rows} columns={columns} />
		</>
	);
}
