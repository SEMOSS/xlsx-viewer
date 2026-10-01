import { Info, OctagonX, TriangleAlert, type LucideIcon } from 'lucide-react';

import { useT, type TranslationKey } from '../../../../i18n';
import {
	VALIDATION_MESSAGE_MAX,
	VALIDATION_TITLE_MAX,
	type ValidationErrorStyle,
} from '../../../commands/data-validation-commands';
import { Checkbox, FIELD, GroupBox, LabeledField, Select } from '../dialog-ui';
import { translatedOptions, type OptionDef, type ValidationTabProps } from './SettingsTab';

const STYLE_OPTIONS: readonly OptionDef<ValidationErrorStyle>[] = [
	{ value: 'stop', labelKey: 'dialog.dataValidation.errorStyle.stop' },
	{ value: 'warning', labelKey: 'dialog.dataValidation.errorStyle.warning' },
	{ value: 'information', labelKey: 'dialog.dataValidation.errorStyle.information' },
];

const STYLE_ICONS: Record<
	ValidationErrorStyle,
	{ Icon: LucideIcon; className: string; labelKey: TranslationKey }
> = {
	stop: {
		Icon: OctagonX,
		className: 'text-destructive',
		labelKey: 'dialog.dataValidation.icon.stop',
	},
	warning: {
		Icon: TriangleAlert,
		className: 'text-amber-500',
		labelKey: 'dialog.dataValidation.icon.warning',
	},
	information: {
		Icon: Info,
		className: 'text-blue-600',
		labelKey: 'dialog.dataValidation.icon.information',
	},
};

/** Error Alert tab: what happens when invalid data is entered. */
export function ErrorAlertTab({ form, onChange }: ValidationTabProps) {
	const t = useT();
	const { Icon, className, labelKey } = STYLE_ICONS[form.errorStyle];
	return (
		<div className='flex flex-col gap-3'>
			<Checkbox
				label={t('dialog.dataValidation.showErrorAlert')}
				checked={form.showErrorMessage}
				onChange={(showErrorMessage) => onChange({ showErrorMessage })}
			/>
			<GroupBox legend={t('dialog.dataValidation.errorAlertGroup')}>
				<div className='grid grid-cols-[130px_minmax(0,1fr)] gap-x-4 gap-y-2 pt-1'>
					<LabeledField label={t('dialog.dataValidation.alertStyle')}>
						<Select
							value={form.errorStyle}
							options={translatedOptions(STYLE_OPTIONS, t)}
							onChange={(errorStyle) => onChange({ errorStyle })}
						/>
					</LabeledField>
					<LabeledField label={t('dialog.dataValidation.titleField')}>
						<input
							className={FIELD}
							maxLength={VALIDATION_TITLE_MAX}
							value={form.errorTitle}
							onChange={(e) => onChange({ errorTitle: e.target.value })}
						/>
					</LabeledField>
					<div className='flex items-start justify-center pt-5'>
						<Icon
							size={44}
							strokeWidth={1.75}
							className={className}
							role='img'
							aria-label={t(labelKey)}
						/>
					</div>
					<LabeledField label={t('dialog.dataValidation.errorMessageField')}>
						<textarea
							className={`${FIELD} h-[112px] resize-none`}
							maxLength={VALIDATION_MESSAGE_MAX}
							value={form.error}
							onChange={(e) => onChange({ error: e.target.value })}
						/>
					</LabeledField>
				</div>
			</GroupBox>
		</div>
	);
}
