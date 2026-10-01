import { useId, useRef, useState } from 'react';
import type { DefinedName } from 'xlsx-viewer-core';
import { definedNameDisplayName, parseRefersTo, validateDefinedName } from 'xlsx-viewer-core';

import { useT } from '../../../i18n';
import type { NameDraft } from '../../commands/name-commands';
import { useViewerStore } from '../../state/use-store';
import { DialogMessage, FIELD, FormDialog, Select, type SelectOption } from './dialog-ui';

export interface NameEditorDialogProps {
	/** `New Name` or `Edit Name`, in the UI language. */
	title: string;
	initial: NameDraft;
	/** The definition being edited: its scope is fixed and it is not its own duplicate. */
	editing?: DefinedName;
	/**
	 * Receives a validated definition (formula without the leading `=`).
	 * Return an error message to keep the dialog open and show it.
	 */
	onSubmit: (def: DefinedName) => string | null | undefined | void;
	onClose: () => void;
}

const WORKBOOK = 'workbook';

/**
 * The New Name / Edit Name dialog: Name, Scope, Comment and Refers to.
 * OK checks the name (Excel's naming rules, unique in its scope) and the
 * formula, and reports problems inline without closing.
 */
export function NameEditorDialog({
	title,
	initial,
	editing,
	onSubmit,
	onClose,
}: NameEditorDialogProps) {
	const t = useT();
	const store = useViewerStore();
	const id = useId();
	const nameRef = useRef<HTMLInputElement>(null);
	const refersRef = useRef<HTMLInputElement>(null);
	const [name, setName] = useState(initial.name);
	const [scope, setScope] = useState(
		initial.scope === undefined ? WORKBOOK : String(initial.scope),
	);
	const [comment, setComment] = useState(initial.comment);
	const [refersTo, setRefersTo] = useState(initial.refersTo);
	const [error, setError] = useState<string | null>(null);

	const scopes: SelectOption<string>[] = [{ value: WORKBOOK, label: t('dialog.names.workbook') }];
	store.workbook.sheets.forEach((sheet, index) => {
		const listed = sheet.state !== 'veryHidden' && sheet.kind !== 'chartsheet';
		if (listed || index === initial.scope) {
			scopes.push({ value: String(index), label: sheet.name });
		}
	});

	const fail = (message: string, field: HTMLInputElement | null) => {
		setError(message);
		field?.focus();
		field?.select();
	};

	const ok = () => {
		const trimmed = name.trim();
		const localSheet = scope === WORKBOOK ? undefined : Number(scope);
		if (!editing || trimmed !== definedNameDisplayName(editing)) {
			const message = validateDefinedName(store.workbook, trimmed, localSheet, editing);
			if (message) {
				fail(message, nameRef.current);
				return;
			}
		}
		let formula: string;
		if (editing && refersTo === initial.refersTo) {
			formula = editing.formula;
		} else {
			const parsed = parseRefersTo(refersTo, store.activeSheet.name);
			if ('error' in parsed) {
				fail(parsed.error, refersRef.current);
				return;
			}
			formula = parsed.formula;
		}
		const def: DefinedName = { name: trimmed, formula };
		if (localSheet !== undefined) {
			def.localSheet = localSheet;
		}
		if (comment.trim() !== '') {
			def.comment = comment;
		}
		const result = onSubmit(def);
		if (typeof result === 'string') {
			setError(result);
		}
	};

	return (
		<FormDialog
			title={title}
			onClose={onClose}
			onOk={ok}
			okDisabled={store.readOnly}
			width={440}
			initialFocusSelector='[data-name-field="name"]'
		>
			<div className='grid grid-cols-[72px_minmax(0,1fr)] items-center gap-x-2 gap-y-2.5'>
				<label htmlFor={`${id}-name`}>{t('dialog.names.nameField')}</label>
				<input
					ref={nameRef}
					id={`${id}-name`}
					data-name-field='name'
					className={FIELD}
					value={name}
					maxLength={255}
					autoComplete='off'
					spellCheck={false}
					onChange={(e) => {
						setName(e.target.value);
						setError(null);
					}}
				/>
				<span aria-hidden='true'>{t('dialog.names.scopeField')}</span>
				<Select
					label={t('dialog.names.scope')}
					value={scope}
					options={scopes}
					disabled={editing !== undefined}
					onChange={setScope}
				/>
				<label htmlFor={`${id}-comment`} className='self-start pt-1'>
					{t('dialog.names.commentField')}
				</label>
				<textarea
					id={`${id}-comment`}
					className={`${FIELD} h-[76px] resize-none`}
					value={comment}
					maxLength={255}
					onChange={(e) => setComment(e.target.value)}
				/>
				<label htmlFor={`${id}-refers`}>{t('dialog.names.refersTo')}</label>
				<input
					ref={refersRef}
					id={`${id}-refers`}
					className={FIELD}
					value={refersTo}
					autoComplete='off'
					spellCheck={false}
					onChange={(e) => {
						setRefersTo(e.target.value);
						setError(null);
					}}
				/>
			</div>
			{error && (
				<DialogMessage>
					<span className='whitespace-pre-line'>{error}</span>
				</DialogMessage>
			)}
		</FormDialog>
	);
}
