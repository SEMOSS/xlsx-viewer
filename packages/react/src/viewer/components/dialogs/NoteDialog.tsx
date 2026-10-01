import { a1, setNote } from 'xlsx-viewer-core';
import { useState } from 'react';

import { useT } from '../../../i18n';
import { useViewerStore } from '../../state/use-store';
import { Dialog, DialogButton } from '../ui/Dialog';

/** Add, edit or delete the note of the active cell. */
export function NoteDialog({ onClose, author }: { onClose: () => void; author: string }) {
	const t = useT();
	const store = useViewerStore();
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const { row, col } = store.selectionOf().active;
	const existing = sheet.notes.find((n) => n.ref.row === row && n.ref.col === col);
	const [text, setText] = useState(existing?.text ?? '');
	const save = (remove = false) => {
		const label = remove ? 'cmd.deleteNote' : existing ? 'cmd.editNote' : 'cmd.newNote';
		store.transact(store.t(label), (tx) => {
			setNote(tx, store.workbook, sheetIndex, row, col, remove ? '' : text, author);
		});
		onClose();
	};
	return (
		<Dialog
			title={t(existing ? 'dialog.note.editTitle' : 'dialog.note.newTitle', { cell: a1(row, col) })}
			onClose={onClose}
			width={380}
			initialFocusSelector='textarea'
			footer={
				<>
					{existing && (
						<DialogButton onClick={() => save(true)} disabled={store.readOnly}>
							{t('dialog.delete')}
						</DialogButton>
					)}
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary onClick={() => save(false)} disabled={store.readOnly}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			<div className='mb-1 text-[12px] font-semibold text-muted-foreground'>
				{existing?.author ?? author}:
			</div>
			<textarea
				className='h-[120px] w-full resize-none rounded-md border border-input bg-[#ffffe1] p-2 text-[13px] text-black outline-none focus:border-primary'
				value={text}
				onChange={(e) => setText(e.target.value)}
				aria-label={t('dialog.note.text')}
			/>
		</Dialog>
	);
}
