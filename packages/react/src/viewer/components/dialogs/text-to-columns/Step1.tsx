import { useId } from 'react';

import { useT, type TranslationKey } from '../../../../i18n';
import { GroupBox } from '../dialog-ui';
import { PREVIEW_LINES, type SplitMode } from './wizard-model';

const TYPES: ReadonlyArray<{
	mode: SplitMode;
	labelKey: TranslationKey;
	descriptionKey: TranslationKey;
}> = [
	{
		mode: 'delimited',
		labelKey: 'dialog.textToColumns.delimited',
		descriptionKey: 'dialog.textToColumns.delimitedHint',
	},
	{
		mode: 'fixed',
		labelKey: 'dialog.textToColumns.fixedWidth',
		descriptionKey: 'dialog.textToColumns.fixedWidthHint',
	},
];

const PREVIEW_BOX =
	'h-[130px] overflow-auto rounded-md border border-input bg-card py-0.5 font-mono text-[12px]';

/** Monospace lines with Excel's row numbers. */
export function LinePreview({ lines, label }: { lines: readonly string[]; label: string }) {
	const id = useId();
	return (
		<div className='flex min-w-0 flex-col gap-1'>
			<span id={id}>{label}</span>
			<div role='region' aria-labelledby={id} tabIndex={0} className={PREVIEW_BOX}>
				{lines.slice(0, PREVIEW_LINES).map((line, i) => (
					<div key={i} className='flex whitespace-pre'>
						<span className='w-8 shrink-0 select-none pr-2 text-right text-muted-foreground'>
							{i + 1}
						</span>
						<span>{line}</span>
					</div>
				))}
			</div>
		</div>
	);
}

/** Step 1: original data type. */
export function Step1({
	mode,
	guessed,
	lines,
	onMode,
}: {
	mode: SplitMode;
	guessed: SplitMode;
	lines: readonly string[];
	onMode: (mode: SplitMode) => void;
}) {
	const t = useT();
	const name = useId();
	return (
		<>
			<p>
				{t(
					guessed === 'fixed'
						? 'dialog.textToColumns.guessedFixed'
						: 'dialog.textToColumns.guessedDelimited',
				)}
				<br />
				{t('dialog.textToColumns.ifCorrect')}
			</p>
			<GroupBox legend={t('dialog.textToColumns.originalType')}>
				<p className='pb-1.5'>{t('dialog.textToColumns.chooseType')}</p>
				<div
					role='radiogroup'
					aria-label={t('dialog.textToColumns.originalType')}
					className='flex flex-col gap-1.5'
				>
					{TYPES.map((type) => (
						<div key={type.mode} className='flex items-center gap-2'>
							<label className='flex w-28 shrink-0 items-center gap-1.5'>
								<input
									type='radio'
									name={name}
									className='accent-[var(--xlsx-primary)]'
									checked={mode === type.mode}
									aria-describedby={`${name}-${type.mode}`}
									onChange={() => onMode(type.mode)}
								/>
								<span>{t(type.labelKey)}</span>
							</label>
							<span id={`${name}-${type.mode}`} className='text-muted-foreground'>
								{t(type.descriptionKey)}
							</span>
						</div>
					))}
				</div>
			</GroupBox>
			<LinePreview lines={lines} label={t('dialog.textToColumns.previewSelected')} />
		</>
	);
}
