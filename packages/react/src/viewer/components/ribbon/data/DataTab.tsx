import {
	ArrowDownAZ,
	ArrowDownZA,
	Columns3,
	Funnel,
	ListChecks,
	ListIndentDecrease,
	ListIndentIncrease,
	Minus,
	Plus,
	Rows3,
} from 'lucide-react';

import { useT } from '../../../../i18n';
import {
	groupSelection,
	toggleDetailAtSelection,
	ungroupSelection,
} from '../../../commands/outline-commands';
import { sortSelection, toggleAutoFilter } from '../../../commands/structure-commands';
import { useShell } from '../../../shell-context';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import { RibbonButton, SplitButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';

/** Data: sort and filter, data tools and outline. */
export function DataTab() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const hasFilter = useStoreSelector(
		(s, st) => (void s.docVersion, Boolean(st.workbook.sheets[s.ui.activeSheet].autoFilter)),
	);
	const readOnly = store.readOnly;
	return (
		<>
			<RibbonGroup label={t('group.sortFilter')}>
				<RibbonButton
					icon={<ArrowDownAZ size={15} />}
					label={t('cmd.sortAZ')}
					disabled={readOnly}
					onClick={() => sortSelection(store, false)}
				/>
				<RibbonButton
					icon={<ArrowDownZA size={15} />}
					label={t('cmd.sortZA')}
					disabled={readOnly}
					onClick={() => sortSelection(store, true)}
				/>
				<RibbonButton
					icon={<Rows3 size={15} />}
					label={t('cmd.sort')}
					showLabel
					disabled={readOnly}
					onClick={() => shell.openDialog('sort')}
				/>
				<RibbonButton
					icon={<Funnel size={15} />}
					label={t('cmd.filter')}
					showLabel
					pressed={hasFilter}
					disabled={readOnly}
					onClick={() => toggleAutoFilter(store)}
					shortcut='Ctrl+Shift+L'
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.dataTools')}>
				<RibbonButton
					icon={<ListChecks size={15} />}
					label={t('cmd.dataValidation')}
					showLabel
					disabled={readOnly}
					onClick={() => shell.openDialog('dataValidation')}
				/>
				<RibbonButton
					icon={<Rows3 size={15} />}
					label={t('cmd.removeDuplicates')}
					showLabel
					disabled={readOnly}
					onClick={() => shell.openDialog('removeDuplicates')}
				/>
				<RibbonButton
					icon={<Columns3 size={15} />}
					label={t('cmd.textToColumns')}
					showLabel
					disabled={readOnly}
					onClick={() => shell.openDialog('textToColumns')}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.outline')}>
				<SplitButton
					icon={<ListIndentIncrease size={15} />}
					label={t('cmd.group')}
					showLabel
					disabled={readOnly}
					shortcut='Alt+Shift+Right'
					onClick={() => groupSelection(store)}
					items={[
						{ label: t('cmd.groupRows'), onSelect: () => groupSelection(store, 'rows') },
						{ label: t('cmd.groupColumns'), onSelect: () => groupSelection(store, 'cols') },
					]}
				/>
				<SplitButton
					icon={<ListIndentDecrease size={15} />}
					label={t('cmd.ungroup')}
					showLabel
					disabled={readOnly}
					shortcut='Alt+Shift+Left'
					onClick={() => ungroupSelection(store)}
					items={[
						{ label: t('cmd.ungroupRows'), onSelect: () => ungroupSelection(store, 'rows') },
						{ label: t('cmd.ungroupColumns'), onSelect: () => ungroupSelection(store, 'cols') },
					]}
				/>
				<RibbonButton
					icon={<Plus size={15} />}
					label={t('cmd.showDetail')}
					disabled={readOnly}
					onClick={() => toggleDetailAtSelection(store, false)}
				/>
				<RibbonButton
					icon={<Minus size={15} />}
					label={t('cmd.hideDetail')}
					disabled={readOnly}
					onClick={() => toggleDetailAtSelection(store, true)}
				/>
			</RibbonGroup>
		</>
	);
}
