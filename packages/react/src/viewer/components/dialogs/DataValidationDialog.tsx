import { useId, useState } from 'react';
import type { DataValidation, RangeAddress } from 'xlsx-viewer-core';

import { useT, type TranslationKey } from '../../../i18n';
import {
	applyValidationForm,
	dataValidationSetup,
	DEFAULT_VALIDATION_FORM,
	type ValidationForm,
} from '../../commands/data-validation-commands';
import { useViewerStore } from '../../state/use-store';
import { DialogButton } from '../ui/Dialog';
import { ErrorAlertTab } from './data-validation/ErrorAlertTab';
import { InputMessageTab } from './data-validation/InputMessageTab';
import { SettingsTab } from './data-validation/SettingsTab';
import { ValidationQuestion } from './data-validation/ValidationQuestion';
import { DialogMessage, FormDialog, TabPanel, TabStrip, type TabDef } from './dialog-ui';
import type { DialogProps } from './types';

type TabId = 'settings' | 'input' | 'error';

const TABS: ReadonlyArray<{ id: TabId; labelKey: TranslationKey }> = [
	{ id: 'settings', labelKey: 'dialog.dataValidation.settings' },
	{ id: 'input', labelKey: 'dialog.dataValidation.inputMessage' },
	{ id: 'error', labelKey: 'dialog.dataValidation.errorAlert' },
];

type Phase = 'mixed' | 'partial' | 'edit';

/**
 * Data > Data Validation: what the selected cells accept, the input message
 * shown when one is selected and the alert shown for invalid entries.
 * Selections with several validations, or with cells without one, first get
 * Excel's question.
 */
export function DataValidationDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const idBase = useId();
	const [setup] = useState(() => dataValidationSetup(store));
	const { state } = setup;
	const [phase, setPhase] = useState<Phase>(
		state.mixed ? 'mixed' : state.partial ? 'partial' : 'edit',
	);
	const [tab, setTab] = useState<TabId>('settings');
	const [form, setForm] = useState<ValidationForm>(setup.form);
	const [targets, setTargets] = useState<RangeAddress[]>(setup.selection);
	/** The validation being edited (for "Apply these changes to all other cells ..."). */
	const [original, setOriginal] = useState<DataValidation | undefined>(
		state.mixed ? undefined : state.rule,
	);
	const [applyToAll, setApplyToAll] = useState(false);
	const [error, setError] = useState<string | null>(null);

	if (phase !== 'edit') {
		return (
			<ValidationQuestion
				kind={phase}
				onCancel={onClose}
				onYes={() => {
					if (phase === 'mixed') {
						setForm(DEFAULT_VALIDATION_FORM);
						setOriginal(undefined);
					}
					setPhase('edit');
				}}
				onNo={() => {
					if (phase === 'mixed') {
						onClose();
						return;
					}
					// Keep the validation on the cells that already have it.
					setTargets(state.covered);
					setPhase('edit');
				}}
			/>
		);
	}

	const patch = (p: Partial<ValidationForm>) => {
		setForm((f) => ({ ...f, ...p }));
		setError(null);
	};

	const ok = () => {
		if (store.readOnly) {
			onClose();
			return;
		}
		const message = applyValidationForm(store, form, targets, applyToAll ? original : undefined);
		if (message) {
			setError(message);
			setTab('settings');
			return;
		}
		onClose();
	};

	const clearAll = () => {
		setForm(DEFAULT_VALIDATION_FORM);
		setApplyToAll(false);
		setError(null);
	};

	const title = t('dialog.dataValidation.title');
	const tabs: TabDef<TabId>[] = TABS.map((d) => ({ id: d.id, label: t(d.labelKey) }));

	return (
		<FormDialog
			title={title}
			onClose={onClose}
			onOk={ok}
			okDisabled={store.readOnly}
			width={460}
			leftButtons={
				<DialogButton onClick={clearAll} disabled={store.readOnly}>
					{t('dialog.dataValidation.clearAll')}
				</DialogButton>
			}
		>
			<TabStrip tabs={tabs} value={tab} onChange={setTab} label={title} idBase={idBase} />
			<TabPanel idBase={idBase} id={tab} className='min-h-[250px]'>
				{tab === 'settings' && (
					<SettingsTab
						form={form}
						onChange={patch}
						applyToAll={applyToAll}
						onApplyToAllChange={setApplyToAll}
						canApplyToAll={original !== undefined}
					/>
				)}
				{tab === 'input' && <InputMessageTab form={form} onChange={patch} />}
				{tab === 'error' && <ErrorAlertTab form={form} onChange={patch} />}
			</TabPanel>
			{error && <DialogMessage>{error}</DialogMessage>}
		</FormDialog>
	);
}
