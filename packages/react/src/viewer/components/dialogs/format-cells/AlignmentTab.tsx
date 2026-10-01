import type { HorizontalAlignment, VerticalAlignment } from 'xlsx-viewer-core';

import { useT, type TranslationKey } from '../../../../i18n';
import {
	MIXED,
	type FormatCellsValues,
	type Mixable,
} from '../../../commands/format-cells-commands';
import { Checkbox, FIELD, FIELD_SM, GroupBox } from '../dialog-ui';
import { degreesToRotation, OrientationDial, rotationToDegrees } from './OrientationDial';

type Patch = Partial<FormatCellsValues>;

interface Option<T extends string> {
	value: T;
	labelKey: TranslationKey;
}

const HORIZONTAL: ReadonlyArray<Option<HorizontalAlignment>> = [
	{ value: 'general', labelKey: 'dialog.formatCells.horizontalAlign.general' },
	{ value: 'left', labelKey: 'dialog.formatCells.horizontalAlign.left' },
	{ value: 'center', labelKey: 'dialog.formatCells.horizontalAlign.center' },
	{ value: 'right', labelKey: 'dialog.formatCells.horizontalAlign.right' },
	{ value: 'fill', labelKey: 'dialog.formatCells.horizontalAlign.fill' },
	{ value: 'justify', labelKey: 'dialog.formatCells.horizontalAlign.justify' },
	{ value: 'centerContinuous', labelKey: 'dialog.formatCells.horizontalAlign.centerContinuous' },
	{ value: 'distributed', labelKey: 'dialog.formatCells.horizontalAlign.distributed' },
];

const VERTICAL: ReadonlyArray<Option<VerticalAlignment>> = [
	{ value: 'top', labelKey: 'dialog.formatCells.verticalAlign.top' },
	{ value: 'center', labelKey: 'dialog.formatCells.verticalAlign.center' },
	{ value: 'bottom', labelKey: 'dialog.formatCells.verticalAlign.bottom' },
	{ value: 'justify', labelKey: 'dialog.formatCells.verticalAlign.justify' },
	{ value: 'distributed', labelKey: 'dialog.formatCells.verticalAlign.distributed' },
];

const DIRECTIONS: ReadonlyArray<Option<'0' | '1' | '2'>> = [
	{ value: '0', labelKey: 'dialog.formatCells.direction.context' },
	{ value: '1', labelKey: 'dialog.formatCells.direction.leftToRight' },
	{ value: '2', labelKey: 'dialog.formatCells.direction.rightToLeft' },
];

const INDENTABLE: ReadonlySet<string> = new Set(['left', 'right', 'distributed']);

function MixedSelect<T extends string>({
	value,
	options,
	onChange,
	disabled,
}: {
	value: Mixable<T>;
	options: ReadonlyArray<Option<T>>;
	onChange: (v: T) => void;
	disabled?: boolean;
}) {
	const t = useT();
	return (
		<select
			className={FIELD}
			value={value === MIXED ? '' : value}
			disabled={disabled}
			onChange={(e) => onChange(e.target.value as T)}
		>
			{value === MIXED && <option value='' />}
			{options.map((o) => (
				<option key={o.value} value={o.value}>
					{t(o.labelKey)}
				</option>
			))}
		</select>
	);
}

function bool(v: Mixable<boolean>): boolean {
	return v === true;
}

/** Format Cells > Alignment. */
export function AlignmentTab({
	draft,
	onChange,
}: {
	draft: FormatCellsValues;
	onChange: (patch: Patch) => void;
}) {
	const t = useT();
	const horizontal = draft.horizontal;
	const canIndent = horizontal !== MIXED && INDENTABLE.has(horizontal);
	const rotation = draft.textRotation;
	const stacked = rotation === 255;
	const degrees = rotation === MIXED || stacked ? null : rotationToDegrees(rotation);
	const setIndent = (n: number) => {
		const indent = Math.max(0, Math.min(250, Math.trunc(n) || 0));
		const patch: Patch = { indent };
		// Indenting a General / Center cell turns it into Left (Indent), as in Excel.
		if (indent > 0 && !canIndent) {
			patch.horizontal = 'left';
		}
		onChange(patch);
	};
	return (
		<div className='flex gap-4'>
			<div className='flex min-w-0 flex-1 flex-col gap-3'>
				<GroupBox legend={t('dialog.formatCells.textAlignment')}>
					<div className='grid grid-cols-[1fr_88px] items-end gap-x-3 gap-y-2'>
						<label className='flex flex-col gap-1'>
							<span>{t('dialog.formatCells.horizontal')}</span>
							<MixedSelect
								value={horizontal}
								options={HORIZONTAL}
								onChange={(h) =>
									onChange({ horizontal: h, indent: INDENTABLE.has(h) ? draft.indent : 0 })
								}
							/>
						</label>
						<label className='flex flex-col gap-1'>
							<span>{t('dialog.formatCells.indent')}</span>
							<input
								type='number'
								min={0}
								max={250}
								className={`${FIELD_SM} w-full`}
								value={draft.indent === MIXED ? '' : draft.indent}
								onChange={(e) => setIndent(Number(e.target.value))}
							/>
						</label>
						<label className='flex flex-col gap-1'>
							<span>{t('dialog.formatCells.vertical')}</span>
							<MixedSelect
								value={draft.vertical}
								options={VERTICAL}
								onChange={(v) => onChange({ vertical: v })}
							/>
						</label>
						<span />
					</div>
					<Checkbox
						className='mt-2'
						label={t('dialog.formatCells.justifyDistributed')}
						disabled={horizontal !== 'distributed'}
						checked={bool(draft.justifyLastLine)}
						indeterminate={draft.justifyLastLine === MIXED}
						onChange={(v) => onChange({ justifyLastLine: v })}
					/>
				</GroupBox>
				<GroupBox legend={t('dialog.formatCells.textControl')}>
					<div className='flex flex-col gap-1.5 pt-1'>
						<Checkbox
							label={t('dialog.formatCells.wrapText')}
							checked={bool(draft.wrapText)}
							indeterminate={draft.wrapText === MIXED}
							onChange={(v) =>
								onChange(v ? { wrapText: true, shrinkToFit: false } : { wrapText: false })
							}
						/>
						<Checkbox
							label={t('dialog.formatCells.shrinkToFit')}
							disabled={draft.wrapText === true}
							checked={bool(draft.shrinkToFit)}
							indeterminate={draft.shrinkToFit === MIXED}
							onChange={(v) => onChange({ shrinkToFit: v })}
						/>
						<Checkbox
							label={t('dialog.formatCells.mergeCells')}
							checked={bool(draft.merged)}
							indeterminate={draft.merged === MIXED}
							onChange={(v) => onChange({ merged: v })}
						/>
					</div>
				</GroupBox>
				<GroupBox legend={t('dialog.formatCells.rightToLeft')}>
					<label className='flex items-center gap-2 pt-1'>
						<span>{t('dialog.formatCells.textDirection')}</span>
						<select
							className={`${FIELD} w-[150px]`}
							value={draft.readingOrder === MIXED ? '' : String(draft.readingOrder)}
							onChange={(e) => onChange({ readingOrder: Number(e.target.value) })}
						>
							{draft.readingOrder === MIXED && <option value='' />}
							{DIRECTIONS.map((d) => (
								<option key={d.value} value={d.value}>
									{t(d.labelKey)}
								</option>
							))}
						</select>
					</label>
				</GroupBox>
			</div>
			<GroupBox legend={t('dialog.formatCells.orientation')} className='w-[176px] shrink-0'>
				<div className='flex items-center gap-2 pt-1'>
					<button
						type='button'
						aria-label={t('dialog.formatCells.verticalText')}
						aria-pressed={stacked}
						className={`flex h-[120px] w-7 flex-col items-center justify-center rounded-sm border text-[12px] leading-[14px] ${stacked ? 'border-primary bg-accent' : 'border-border hover:bg-accent'}`}
						onClick={() => onChange({ textRotation: stacked ? 0 : 255 })}
					>
						{Array.from(t('dialog.formatCells.previewText')).map((ch, i) => (
							<span key={i}>{ch}</span>
						))}
					</button>
					<OrientationDial
						degrees={degrees}
						onChange={(d) => onChange({ textRotation: degreesToRotation(d) })}
					/>
				</div>
				<label className='mt-2 flex items-center justify-center gap-2'>
					<input
						type='number'
						min={-90}
						max={90}
						aria-label={t('dialog.formatCells.degrees')}
						className={`${FIELD_SM} w-16`}
						disabled={stacked}
						value={degrees === null ? '' : degrees}
						onChange={(e) =>
							onChange({ textRotation: degreesToRotation(Number(e.target.value) || 0) })
						}
					/>
					<span>{t('dialog.formatCells.degrees')}</span>
				</label>
			</GroupBox>
		</div>
	);
}
