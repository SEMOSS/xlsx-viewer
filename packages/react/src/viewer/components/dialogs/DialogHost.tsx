import type { DialogKind } from '../../shell-context';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { FindReplaceDialog } from './FindReplaceDialog';
import { NoteDialog } from './NoteDialog';
import { DIALOG_REGISTRY } from './registry';
import { GoToDialog, RenameSheetDialog, SizeDialog, ZoomDialog } from './SimpleDialogs';

export type { DialogProps } from './types';

/** Renders the dialog requested in the UI state. */
export function DialogHost({ author }: { author: string }) {
	const store = useViewerStore();
	const dialog = useStoreSelector((s) => s.ui.dialog);
	if (!dialog) {
		return null;
	}
	const close = () => store.setUi({ dialog: null });
	switch (dialog.kind as DialogKind) {
		case 'find':
			return <FindReplaceDialog mode='find' onClose={close} />;
		case 'replace':
			return <FindReplaceDialog mode='replace' onClose={close} />;
		case 'goTo':
			return <GoToDialog onClose={close} />;
		case 'rowHeight':
			return <SizeDialog kind='rowHeight' onClose={close} />;
		case 'columnWidth':
			return <SizeDialog kind='columnWidth' onClose={close} />;
		case 'zoom':
			return <ZoomDialog onClose={close} />;
		case 'renameSheet':
			return <RenameSheetDialog onClose={close} />;
		case 'note':
			return <NoteDialog onClose={close} author={author} />;
		default: {
			const Extra = DIALOG_REGISTRY[dialog.kind as DialogKind];
			if (Extra) {
				return <Extra onClose={close} payload={dialog.payload} author={author} />;
			}
			return null;
		}
	}
}

/** Dialog kinds that have an implementation (others report "not available yet"). */
export function isDialogAvailable(kind: DialogKind): boolean {
	return (
		['find', 'replace', 'goTo', 'rowHeight', 'columnWidth', 'zoom', 'renameSheet', 'note'].includes(
			kind,
		) || DIALOG_REGISTRY[kind] !== undefined
	);
}
