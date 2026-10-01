import { ClipboardPaste, Copy, Paintbrush, Scissors } from 'lucide-react';

import { useT } from '../../../../i18n';
import { copyToSystem, pasteFromSystem } from '../../../commands/clipboard-commands';
import { currentRange } from '../../../../engine/selection/selection';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import { RibbonButton, SplitButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';

export function ClipboardGroup() {
	const t = useT();
	const store = useViewerStore();
	const painter = useStoreSelector((s) => s.ui.formatPainter);
	const readOnly = store.readOnly;
	return (
		<RibbonGroup label={t('group.clipboard')}>
			<SplitButton
				icon={<ClipboardPaste size={16} />}
				label={t('cmd.paste')}
				showLabel
				disabled={readOnly}
				shortcut='Ctrl+V'
				onClick={() => void pasteFromSystem(store)}
				items={[
					{ label: t('cmd.paste'), onSelect: () => void pasteFromSystem(store, 'all') },
					{ label: t('cmd.pasteValues'), onSelect: () => void pasteFromSystem(store, 'values') },
					{
						label: t('cmd.pasteFormulas'),
						onSelect: () => void pasteFromSystem(store, 'formulas'),
					},
					{
						label: t('cmd.pasteFormatting'),
						onSelect: () => void pasteFromSystem(store, 'formats'),
					},
					{
						label: t('cmd.pasteTranspose'),
						onSelect: () => void pasteFromSystem(store, 'transpose'),
					},
				]}
			/>
			<RibbonButton
				icon={<Scissors size={15} />}
				label={t('cmd.cut')}
				shortcut='Ctrl+X'
				disabled={readOnly}
				onClick={() => void copyToSystem(store, true)}
			/>
			<RibbonButton
				icon={<Copy size={15} />}
				label={t('cmd.copy')}
				shortcut='Ctrl+C'
				onClick={() => void copyToSystem(store, false)}
			/>
			<RibbonButton
				icon={<Paintbrush size={15} />}
				label={t('cmd.formatPainter')}
				pressed={painter !== null}
				disabled={readOnly}
				onClick={() =>
					store.setUi({
						formatPainter: painter
							? null
							: {
									mode: 'once',
									sheet: store.state.activeSheet,
									range: currentRange(store.selectionOf()),
								},
					})
				}
				onDoubleClick={() =>
					store.setUi({
						formatPainter: {
							mode: 'sticky',
							sheet: store.state.activeSheet,
							range: currentRange(store.selectionOf()),
						},
					})
				}
			/>
		</RibbonGroup>
	);
}
