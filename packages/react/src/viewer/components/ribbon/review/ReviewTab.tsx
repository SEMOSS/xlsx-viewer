import {
	ChevronLeft,
	ChevronRight,
	Eye,
	Lock,
	LockOpen,
	MessageSquarePlus,
	Search,
	ShieldCheck,
	StickyNote,
	Trash2,
} from 'lucide-react';

import { useT } from '../../../../i18n';
import {
	deleteNotesCmd,
	goToAdjacentNote,
	toggleActiveNoteVisible,
	toggleAllNotes,
} from '../../../commands/note-commands';
import {
	toggleSheetProtection,
	toggleWorkbookProtection,
} from '../../../commands/protection-commands';
import { useShell } from '../../../shell-context';
import { shallowEqual, useStoreSelector, useViewerStore } from '../../../state/use-store';
import { RibbonButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';

/** Review: notes, protection and find. */
export function ReviewTab() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const state = useStoreSelector((snap, st) => {
		void snap.docVersion;
		const sheet = st.workbook.sheets[snap.ui.activeSheet];
		const { row, col } = st.selectionOf(snap.ui.activeSheet).active;
		const note = sheet.notes.find((n) => n.ref.row === row && n.ref.col === col);
		return {
			hasNote: !!note,
			noteVisible: !!note?.visible,
			anyNotes: sheet.notes.length > 0,
			allVisible: sheet.notes.length > 0 && sheet.notes.every((n) => n.visible),
			sheetProtected: sheet.protection?.sheet === true,
			workbookProtected: st.workbook.protection?.lockStructure === true,
		};
	}, shallowEqual);
	const ro = store.readOnly;
	return (
		<>
			<RibbonGroup label={t('group.notes')}>
				<RibbonButton
					icon={<MessageSquarePlus size={15} />}
					label={state.hasNote ? t('cmd.editNote') : t('cmd.newNote')}
					showLabel
					disabled={ro}
					onClick={() => shell.openDialog('note')}
				/>
				<RibbonButton
					icon={<Trash2 size={15} />}
					label={t('cmd.deleteNote')}
					disabled={ro || !state.hasNote}
					onClick={() => deleteNotesCmd(store)}
				/>
				<RibbonButton
					icon={<ChevronLeft size={15} />}
					label={t('cmd.previousNote')}
					disabled={!state.anyNotes}
					onClick={() => goToAdjacentNote(store, -1)}
				/>
				<RibbonButton
					icon={<ChevronRight size={15} />}
					label={t('cmd.nextNote')}
					disabled={!state.anyNotes}
					onClick={() => goToAdjacentNote(store, 1)}
				/>
				<RibbonButton
					icon={<StickyNote size={15} />}
					label={t('cmd.showHideNote')}
					pressed={state.noteVisible}
					disabled={ro || !state.hasNote}
					onClick={() => toggleActiveNoteVisible(store)}
				/>
				<RibbonButton
					icon={<Eye size={15} />}
					label={t('cmd.showAllNotes')}
					showLabel
					pressed={state.allVisible}
					disabled={ro || !state.anyNotes}
					onClick={() => toggleAllNotes(store)}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.protect')}>
				<RibbonButton
					icon={state.sheetProtected ? <LockOpen size={15} /> : <Lock size={15} />}
					label={state.sheetProtected ? t('protect.unprotectSheet') : t('protect.protectSheet')}
					showLabel
					disabled={ro}
					onClick={() => toggleSheetProtection(store, shell)}
				/>
				<RibbonButton
					icon={<ShieldCheck size={15} />}
					label={t('protect.protectWorkbook')}
					showLabel
					pressed={state.workbookProtected}
					disabled={ro}
					onClick={() => toggleWorkbookProtection(store, shell)}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.editing')}>
				<RibbonButton
					icon={<Search size={15} />}
					label={t('cmd.find')}
					showLabel
					onClick={() => shell.openDialog('find')}
				/>
			</RibbonGroup>
		</>
	);
}
