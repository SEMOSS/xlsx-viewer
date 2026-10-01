import { Check, X } from 'lucide-react';
import {
	useEffect,
	useId,
	useMemo,
	useRef,
	useState,
	type KeyboardEvent,
	type ReactNode,
} from 'react';
import type { DefinedName } from 'xlsx-viewer-core';
import { parseRefersTo } from 'xlsx-viewer-core';

import { useT } from '../../../i18n';
import {
	defineNameCmd,
	definedNameKey,
	deleteNameCmd,
	editNameCmd,
	editNameDraft,
	nameManagerRows,
	newNameDraft,
	type NameDraft,
	type NameRow,
} from '../../commands/name-commands';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { Dialog, DialogButton } from '../ui/Dialog';
import { DialogMessage, FIELD } from './dialog-ui';
import { NameEditorDialog } from './NameEditorDialog';
import { ConfirmBar, type ConfirmRequest } from './names/ConfirmBar';
import { NameList } from './names/NameList';
import type { DialogProps } from './types';

interface EditorState {
	draft: NameDraft;
	/** The definition being edited (absent for New Name). */
	target?: DefinedName;
}

const ICON_BUTTON =
	'inline-flex h-[26px] w-[26px] shrink-0 items-center justify-center rounded-md border ' +
	'border-border bg-card hover:bg-accent disabled:opacity-40';

function IconButton(props: {
	label: string;
	title: string;
	disabled: boolean;
	onClick: () => void;
	children: ReactNode;
}) {
	return (
		<button
			type='button'
			aria-label={props.label}
			title={props.title}
			disabled={props.disabled}
			onClick={props.onClick}
			className={ICON_BUTTON}
		>
			{props.children}
		</button>
	);
}

/**
 * Formulas > Name Manager: every visible defined name with its value,
 * reference, scope and comment. New / Edit open the name editor on top,
 * Delete asks first, and the Refers to box edits the selected name in
 * place. Each change is its own undoable step.
 */
export function NameManagerDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const version = useStoreSelector((s) => s.docVersion);
	// `t` too: the Workbook scope follows the UI language.
	// eslint-disable-next-line react-hooks/exhaustive-deps
	const rows = useMemo(() => nameManagerRows(store), [store, version, t]);
	const [selected, setSelected] = useState<string | null>(() => rows[0]?.key ?? null);
	/** Edited Refers to text, or null while it shows the selected name's reference. */
	const [draft, setDraft] = useState<string | null>(null);
	const [error, setError] = useState<string | null>(null);
	const [editor, setEditor] = useState<EditorState | null>(null);
	const [confirm, setConfirm] = useState<ConfirmRequest | null>(null);
	const bodyRef = useRef<HTMLDivElement>(null);
	const asked = useRef(false);
	const refId = useId();
	const readOnly = store.readOnly;
	const current = rows.find((r) => r.key === selected) ?? null;
	const dirty = draft !== null;

	useEffect(() => {
		// Give the focus back to the list once a question is answered.
		if (confirm) {
			asked.current = true;
		} else if (asked.current) {
			asked.current = false;
			if (!document.activeElement || document.activeElement === document.body) {
				bodyRef.current?.querySelector<HTMLElement>('[role="grid"]')?.focus();
			}
		}
	}, [confirm]);

	const discard = () => {
		setDraft(null);
		setError(null);
	};

	const commit = (): boolean => {
		if (!current || draft === null) {
			return true;
		}
		const parsed = parseRefersTo(draft, store.activeSheet.name);
		if ('error' in parsed) {
			setError(parsed.error);
			return false;
		}
		const result = editNameCmd(store, current.def, { ...current.def, formula: parsed.formula });
		if (result.error) {
			setError(result.error);
			return false;
		}
		discard();
		if (result.def) {
			setSelected(definedNameKey(result.def));
		}
		return true;
	};

	/** Run `then`, first asking whether to keep an edited Refers to (like Excel). */
	const settle = (then: () => void) => {
		if (!dirty || readOnly) {
			discard();
			then();
			return;
		}
		setConfirm({
			message: t('dialog.names.saveReference'),
			okLabel: t('dialog.yes'),
			cancelLabel: t('dialog.no'),
			onOk: () => commit() && then(),
			onCancel: () => {
				discard();
				then();
			},
		});
	};

	const close = () => settle(onClose);
	const select = (key: string) => key !== selected && settle(() => setSelected(key));
	const openNew = () => settle(() => setEditor({ draft: newNameDraft(store) }));
	const openEdit = (row: NameRow | null) => {
		if (row && !readOnly) {
			settle(() => {
				setSelected(row.key);
				setEditor({ draft: editNameDraft(row.def), target: row.def });
			});
		}
	};
	const askDelete = () => {
		const row = current;
		if (!row || readOnly) {
			return;
		}
		setConfirm({
			message: t('dialog.names.confirmDelete', { name: row.name }),
			okLabel: t('dialog.ok'),
			cancelLabel: t('dialog.cancel'),
			onOk: () => {
				const at = rows.findIndex((r) => r.key === row.key);
				deleteNameCmd(store, row.def);
				const next = nameManagerRows(store);
				setSelected(next[Math.min(at, next.length - 1)]?.key ?? null);
				discard();
			},
		});
	};

	const submit = (def: DefinedName): string | null => {
		const target = editor?.target;
		const result = target ? editNameCmd(store, target, def) : defineNameCmd(store, def);
		if (result.error) {
			return result.error;
		}
		if (result.def) {
			setSelected(definedNameKey(result.def));
		}
		setEditor(null);
		return null;
	};

	/** Enter means Close (the default button), except on buttons and in the Refers to box. */
	const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
		const target = e.target as HTMLElement;
		if (
			e.key !== 'Enter' ||
			e.shiftKey ||
			e.altKey ||
			e.ctrlKey ||
			e.metaKey ||
			e.nativeEvent.isComposing ||
			target.closest('[data-enter="self"]') ||
			['BUTTON', 'TEXTAREA', 'A'].includes(target.tagName)
		) {
			return;
		}
		e.preventDefault();
		if (target.id === refId) {
			if (!readOnly) {
				commit();
			}
			return;
		}
		close();
	};

	return (
		<>
			<Dialog
				title={t('dialog.names.managerTitle')}
				onClose={close}
				width={780}
				initialFocusSelector='[role="grid"]'
				footer={
					<DialogButton primary onClick={close}>
						{t('dialog.close')}
					</DialogButton>
				}
			>
				<div ref={bodyRef} className='flex flex-col gap-3' onKeyDown={onKeyDown}>
					<div className='flex flex-col gap-3' inert={confirm !== null}>
						<div className='flex gap-2'>
							<DialogButton onClick={openNew} disabled={readOnly}>
								{t('dialog.names.new')}
							</DialogButton>
							<DialogButton onClick={() => openEdit(current)} disabled={readOnly || !current}>
								{t('dialog.names.edit')}
							</DialogButton>
							<DialogButton onClick={askDelete} disabled={readOnly || !current}>
								{t('dialog.delete')}
							</DialogButton>
						</div>
						<NameList
							rows={rows}
							selected={selected}
							onSelect={select}
							onActivate={(key) => openEdit(rows.find((r) => r.key === key) ?? null)}
							onDelete={askDelete}
						/>
						<div className='flex flex-col gap-1'>
							<label htmlFor={refId}>{t('dialog.names.refersTo')}</label>
							<div className='flex items-center gap-1.5'>
								<IconButton
									label={t('dialog.names.cancelEdit')}
									title={t('dialog.cancel')}
									disabled={readOnly || !dirty}
									onClick={discard}
								>
									<X size={14} aria-hidden='true' />
								</IconButton>
								<IconButton
									label={t('dialog.names.commitEdit')}
									title={t('dialog.names.commit')}
									disabled={readOnly || !dirty}
									onClick={commit}
								>
									<Check size={14} aria-hidden='true' />
								</IconButton>
								<input
									id={refId}
									className={FIELD}
									value={draft ?? current?.refersTo ?? ''}
									readOnly={readOnly}
									disabled={!current}
									autoComplete='off'
									spellCheck={false}
									onChange={(e) => {
										setDraft(e.target.value);
										setError(null);
									}}
								/>
							</div>
						</div>
					</div>
					{error && <DialogMessage>{error}</DialogMessage>}
					{confirm && <ConfirmBar request={confirm} onDone={() => setConfirm(null)} />}
				</div>
			</Dialog>
			{editor && (
				<NameEditorDialog
					title={t(editor.target ? 'dialog.names.editTitle' : 'dialog.names.newTitle')}
					initial={editor.draft}
					editing={editor.target}
					onSubmit={submit}
					onClose={() => setEditor(null)}
				/>
			)}
		</>
	);
}
