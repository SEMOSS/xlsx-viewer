import { useState } from 'react';

import { useT } from '../../../i18n';
import { hideSheetCmd } from '../../commands/sheet-commands';
import { useViewerStore } from '../../state/use-store';
import { FormDialog, ListBox } from './dialog-ui';
import type { DialogProps } from './types';

/** Home > Format > Hide & Unhide > Unhide Sheet: pick one hidden sheet to show again. */
export function UnhideSheetDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	// Very hidden sheets are not offered, like Excel.
	const hidden = store.workbook.sheets
		.map((sheet, index) => ({ sheet, index }))
		.filter((s) => s.sheet.state === 'hidden');
	const [selected, setSelected] = useState<string | null>(
		hidden[0] ? String(hidden[0].index) : null,
	);

	const ok = () => {
		if (selected === null) {
			onClose();
			return;
		}
		onClose();
		hideSheetCmd(store, Number(selected), false);
	};

	return (
		<FormDialog
			title={t('dialog.unhideSheet.title')}
			onClose={onClose}
			onOk={ok}
			okDisabled={selected === null || store.readOnly}
			width={300}
			initialFocusSelector='[role="listbox"]'
		>
			<span id='xv-unhide-label'>{t('dialog.unhideSheet.sheetsLabel')}</span>
			<ListBox
				label={t('dialog.unhideSheet.sheets')}
				items={hidden}
				getKey={(s) => String(s.index)}
				getText={(s) => s.sheet.name}
				render={(s) => s.sheet.name}
				value={selected}
				onChange={setSelected}
				onActivate={(key) => {
					setSelected(key);
					onClose();
					hideSheetCmd(store, Number(key), false);
				}}
				className='h-[150px]'
				empty={
					<div className='px-2 py-1 text-muted-foreground'>{t('dialog.unhideSheet.noHidden')}</div>
				}
			/>
		</FormDialog>
	);
}
