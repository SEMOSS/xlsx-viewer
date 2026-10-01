import type { DataValidationOperator, DataValidationType } from 'xlsx-viewer-core';
import { validationUsesOperator, validationValueLabels } from 'xlsx-viewer-core';

import { useT, type Translate, type TranslationKey } from '../../../../i18n';
import {
	validationValueCaptions,
	type ValidationForm,
} from '../../../commands/data-validation-commands';
import { Checkbox, FIELD, GroupBox, LabeledField, Select, type SelectOption } from '../dialog-ui';

/** A drop-down option whose caption is translated at render time. */
export interface OptionDef<T extends string> {
	value: T;
	labelKey: TranslationKey;
}

const ALLOW_OPTIONS: readonly OptionDef<DataValidationType>[] = [
	{ value: 'none', labelKey: 'dialog.dataValidation.type.none' },
	{ value: 'whole', labelKey: 'dialog.dataValidation.type.whole' },
	{ value: 'decimal', labelKey: 'dialog.dataValidation.type.decimal' },
	{ value: 'list', labelKey: 'dialog.dataValidation.type.list' },
	{ value: 'date', labelKey: 'dialog.dataValidation.type.date' },
	{ value: 'time', labelKey: 'dialog.dataValidation.type.time' },
	{ value: 'textLength', labelKey: 'dialog.dataValidation.type.textLength' },
	{ value: 'custom', labelKey: 'dialog.dataValidation.type.custom' },
];

const DATA_OPTIONS: readonly OptionDef<DataValidationOperator>[] = [
	{ value: 'between', labelKey: 'dialog.dataValidation.operator.between' },
	{ value: 'notBetween', labelKey: 'dialog.dataValidation.operator.notBetween' },
	{ value: 'equal', labelKey: 'dialog.dataValidation.operator.equal' },
	{ value: 'notEqual', labelKey: 'dialog.dataValidation.operator.notEqual' },
	{ value: 'greaterThan', labelKey: 'dialog.dataValidation.operator.greaterThan' },
	{ value: 'lessThan', labelKey: 'dialog.dataValidation.operator.lessThan' },
	{ value: 'greaterThanOrEqual', labelKey: 'dialog.dataValidation.operator.greaterThanOrEqual' },
	{ value: 'lessThanOrEqual', labelKey: 'dialog.dataValidation.operator.lessThanOrEqual' },
];

/** Drop-down options with their captions in the UI language. */
export function translatedOptions<T extends string>(
	defs: readonly OptionDef<T>[],
	t: Translate,
): SelectOption<T>[] {
	return defs.map((d) => ({ value: d.value, label: t(d.labelKey) }));
}

export interface ValidationTabProps {
	form: ValidationForm;
	onChange: (patch: Partial<ValidationForm>) => void;
}

export interface SettingsTabProps extends ValidationTabProps {
	applyToAll: boolean;
	onApplyToAllChange: (value: boolean) => void;
	/** The selection has a validation whose other cells can be changed along. */
	canApplyToAll: boolean;
}

/** Settings tab: what the cells allow ("Validation criteria"). */
export function SettingsTab({
	form,
	onChange,
	applyToAll,
	onApplyToAllChange,
	canApplyToAll,
}: SettingsTabProps) {
	const t = useT();
	const labels = validationValueLabels(form.type, form.operator);
	const captions = validationValueCaptions(form.type, form.operator, t);
	const hasData = validationUsesOperator(form.type);
	const isList = form.type === 'list';
	return (
		<div className='flex flex-col gap-3'>
			<GroupBox legend={t('dialog.dataValidation.criteria')}>
				<div className='grid grid-cols-[minmax(0,1fr)_150px] items-end gap-x-4 gap-y-2 pt-1'>
					<LabeledField label={t('dialog.dataValidation.allow')}>
						<Select
							value={form.type}
							options={translatedOptions(ALLOW_OPTIONS, t)}
							onChange={(type) => onChange({ type, value1: '', value2: '' })}
						/>
					</LabeledField>
					<Checkbox
						className='pb-1'
						label={t('dialog.dataValidation.ignoreBlank')}
						checked={form.ignoreBlank}
						disabled={form.type === 'none'}
						onChange={(ignoreBlank) => onChange({ ignoreBlank })}
					/>
					{(hasData || isList) && (
						<>
							{hasData ? (
								<LabeledField label={t('dialog.dataValidation.data')}>
									<Select
										value={form.operator}
										options={translatedOptions(DATA_OPTIONS, t)}
										onChange={(operator) => onChange({ operator })}
									/>
								</LabeledField>
							) : (
								<span />
							)}
							{isList ? (
								<Checkbox
									className='pb-1'
									label={t('dialog.dataValidation.inCellDropdown')}
									checked={form.inCellDropdown}
									onChange={(inCellDropdown) => onChange({ inCellDropdown })}
								/>
							) : (
								<span />
							)}
						</>
					)}
					{labels.map((label, i) => (
						<LabeledField key={`${form.type}-${label}`} label={captions[i]} className='col-span-2'>
							<input
								className={FIELD}
								value={i === 0 ? form.value1 : form.value2}
								spellCheck={false}
								autoComplete='off'
								onChange={(e) =>
									onChange(i === 0 ? { value1: e.target.value } : { value2: e.target.value })
								}
							/>
						</LabeledField>
					))}
				</div>
			</GroupBox>
			<Checkbox
				label={t('dialog.dataValidation.applyToAll')}
				checked={applyToAll}
				disabled={!canApplyToAll}
				onChange={onApplyToAllChange}
			/>
		</div>
	);
}
