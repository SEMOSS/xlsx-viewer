import { useState } from 'react';

import { useT } from '../../../i18n';
import { defineNameCmd, newNameDraft } from '../../commands/name-commands';
import { useViewerStore } from '../../state/use-store';
import { NameEditorDialog } from './NameEditorDialog';
import type { DialogProps } from './types';

/**
 * Formulas > Define Name: the New Name dialog, prefilled with a name
 * suggested from the selection's label, Workbook scope and the selection
 * as Refers to. OK adds the name as one undoable step.
 */
export function DefineNameDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const [initial] = useState(() => newNameDraft(store));
	return (
		<NameEditorDialog
			title={t('dialog.names.newTitle')}
			initial={initial}
			onClose={onClose}
			onSubmit={(def) => {
				const { error } = defineNameCmd(store, def);
				if (error) {
					return error;
				}
				onClose();
				return null;
			}}
		/>
	);
}
