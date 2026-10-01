import { useT } from '../../../../i18n';
import {
	VALIDATION_MESSAGE_MAX,
	VALIDATION_TITLE_MAX,
} from '../../../commands/data-validation-commands';
import { Checkbox, FIELD, GroupBox, LabeledField } from '../dialog-ui';
import type { ValidationTabProps } from './SettingsTab';

/** Input Message tab: the prompt shown when a validated cell is selected. */
export function InputMessageTab({ form, onChange }: ValidationTabProps) {
	const t = useT();
	return (
		<div className='flex flex-col gap-3'>
			<Checkbox
				label={t('dialog.dataValidation.showInputMessage')}
				checked={form.showInputMessage}
				onChange={(showInputMessage) => onChange({ showInputMessage })}
			/>
			<GroupBox legend={t('dialog.dataValidation.inputMessageGroup')}>
				<div className='flex flex-col gap-2 pt-1'>
					<LabeledField label={t('dialog.dataValidation.titleField')}>
						<input
							className={FIELD}
							maxLength={VALIDATION_TITLE_MAX}
							value={form.promptTitle}
							onChange={(e) => onChange({ promptTitle: e.target.value })}
						/>
					</LabeledField>
					<LabeledField label={t('dialog.dataValidation.inputMessageField')}>
						<textarea
							className={`${FIELD} h-[112px] resize-none`}
							maxLength={VALIDATION_MESSAGE_MAX}
							value={form.prompt}
							onChange={(e) => onChange({ prompt: e.target.value })}
						/>
					</LabeledField>
				</div>
			</GroupBox>
		</div>
	);
}
