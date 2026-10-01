import { useId, useState } from 'react';

import { useT } from '../../../i18n';
import { removeHyperlinksCmd } from '../../commands/hyperlink-commands';
import {
	applyHyperlinkForm,
	defaultLinkText,
	hyperlinkDialogSetup,
	type HyperlinkForm,
	type LinkKind,
} from '../../commands/hyperlink-dialog-commands';
import { useViewerStore } from '../../state/use-store';
import { DialogButton } from '../ui/Dialog';
import { DialogMessage, FIELD, FormDialog, LabeledField, TabPanel } from './dialog-ui';
import { LinkTypeTabs } from './hyperlink/LinkTypeTabs';
import { PlaceList } from './hyperlink/PlaceList';
import type { DialogProps } from './types';

/** Field focused when the dialog opens, per link type. */
const FIRST_FIELD: Record<LinkKind, string> = {
	file: '[data-hl-field="address"]',
	place: '[data-hl-field="cellref"]',
	email: '[data-hl-field="email"]',
};

const TOGGLE_BUTTON =
	'min-w-[76px] shrink-0 rounded-md border border-border bg-card px-3 py-1.5 text-[13px] ' +
	'hover:bg-accent';

/**
 * Insert > Link (Ctrl+K): link the selected cells to a web page or file,
 * a place in this workbook (a cell on a sheet or a defined name) or an
 * e-mail address. Shown as "Edit Hyperlink" with Remove Link when the
 * active cell already has a link. OK applies one undoable step.
 */
export function HyperlinkDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const id = useId();
	const [setup] = useState(() => hyperlinkDialogSetup(store));
	const [form, setForm] = useState<HyperlinkForm>(setup.form);
	const [textTouched, setTextTouched] = useState(false);
	const [showTip, setShowTip] = useState(setup.form.tooltip !== '');
	const [error, setError] = useState<string | null>(null);
	const readOnly = store.readOnly;
	const several = form.text === null;
	const place = setup.places.find((p) => p.key === form.place);
	// Shown, disabled, in Text to display when several cells hold values (never written).
	let shownText = form.text ?? t('dialog.hyperlink.selectionInDocument');
	if (!several && setup.mirror && !textTouched) {
		shownText = defaultLinkText(form, setup.places);
	}

	const update = (patch: Partial<HyperlinkForm>) => {
		setForm((f) => ({ ...f, ...patch }));
		setError(null);
	};

	const ok = () => {
		const message = applyHyperlinkForm(store, setup, { ...form, text: several ? null : shownText });
		if (message) {
			setError(message);
			return;
		}
		onClose();
	};

	const removeLink = () => {
		removeHyperlinksCmd(store);
		onClose();
	};

	const input = (
		field: keyof HyperlinkForm,
		extra: { disabled?: boolean; marker?: string } = {},
	) => (
		<input
			className={FIELD}
			data-hl-field={extra.marker}
			value={String(form[field] ?? '')}
			disabled={extra.disabled}
			autoComplete='off'
			spellCheck={false}
			onChange={(e) => update({ [field]: e.target.value })}
		/>
	);

	return (
		<FormDialog
			title={t(setup.editing ? 'dialog.hyperlink.editTitle' : 'dialog.hyperlink.insertTitle')}
			onClose={onClose}
			onOk={ok}
			okDisabled={readOnly}
			width={640}
			initialFocusSelector={FIRST_FIELD[setup.form.kind]}
			leftButtons={
				setup.editing ? (
					<DialogButton onClick={removeLink} disabled={readOnly}>
						{t('dialog.hyperlink.removeLink')}
					</DialogButton>
				) : undefined
			}
		>
			<div className='flex gap-3'>
				<div className='flex shrink-0 flex-col gap-1'>
					<span id={`${id}-link-to`}>{t('dialog.hyperlink.linkTo')}</span>
					<LinkTypeTabs
						value={form.kind}
						onChange={(kind) => update({ kind })}
						idBase={id}
						labelledBy={`${id}-link-to`}
					/>
				</div>
				<div className='flex min-w-0 flex-1 flex-col gap-3'>
					<div className='flex items-end gap-2'>
						<LabeledField label={t('dialog.hyperlink.textToDisplay')} className='flex-1'>
							<input
								className={FIELD}
								value={shownText}
								disabled={several}
								autoComplete='off'
								onChange={(e) => {
									setTextTouched(true);
									update({ text: e.target.value });
								}}
							/>
						</LabeledField>
						<button
							type='button'
							className={TOGGLE_BUTTON}
							aria-expanded={showTip}
							aria-controls={`${id}-tip`}
							onClick={() => setShowTip((v) => !v)}
						>
							{t('dialog.hyperlink.screenTip')}
						</button>
					</div>
					{showTip && (
						<div id={`${id}-tip`}>
							<LabeledField label={t('dialog.hyperlink.screenTipText')}>
								<input
									className={FIELD}
									value={form.tooltip}
									maxLength={255}
									autoComplete='off'
									onChange={(e) => update({ tooltip: e.target.value })}
								/>
							</LabeledField>
						</div>
					)}
					<TabPanel idBase={id} id={form.kind} className='flex min-h-[236px] flex-col gap-2.5'>
						{form.kind === 'file' && (
							<LabeledField label={t('dialog.hyperlink.address')}>
								{input('address', { marker: 'address' })}
							</LabeledField>
						)}
						{form.kind === 'place' && (
							<>
								<LabeledField label={t('dialog.hyperlink.cellReference')}>
									{input('cellRef', { marker: 'cellref', disabled: place?.group === 'name' })}
								</LabeledField>
								<span aria-hidden='true'>{t('dialog.hyperlink.selectPlace')}</span>
								<PlaceList
									places={setup.places}
									value={form.place}
									onChange={(key) => update({ place: key })}
									label={t('dialog.hyperlink.selectPlace')}
								/>
							</>
						)}
						{form.kind === 'email' && (
							<>
								<LabeledField label={t('dialog.hyperlink.emailField')}>
									{input('email', { marker: 'email' })}
								</LabeledField>
								<LabeledField label={t('dialog.hyperlink.subject')}>
									{input('subject')}
								</LabeledField>
							</>
						)}
					</TabPanel>
					{error && <DialogMessage>{error}</DialogMessage>}
				</div>
			</div>
		</FormDialog>
	);
}
