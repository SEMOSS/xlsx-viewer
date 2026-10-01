import {
	ArrowDownAZ,
	ArrowDownToLine,
	ArrowDownZA,
	ArrowRightToLine,
	Eraser,
	Funnel,
	Search,
	Sigma,
	SquarePen,
	TableProperties,
	Trash2,
	Plus,
} from 'lucide-react';

import { useT } from '../../../../i18n';
import { clearAll } from '../../../commands/format-commands';
import { fillDownOrRight } from '../../../commands/fill-commands';
import {
	autoFitColumns,
	autoFitRows,
	hideColumns,
	hideRows,
} from '../../../commands/layout-commands';
import { autoSum } from '../../../commands/misc-commands';
import { deleteSheetCmd, hideSheetCmd, insertSheetCmd } from '../../../commands/sheet-commands';
import {
	deleteCellsCmd,
	insertCellsCmd,
	sortSelection,
	toggleAutoFilter,
} from '../../../commands/structure-commands';
import { useViewerStore } from '../../../state/use-store';
import { useShell } from '../../../shell-context';
import { MenuButton, SplitButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';

export function CellsGroup() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const readOnly = store.readOnly;
	const active = () => store.selectionOf().active;
	return (
		<RibbonGroup label={t('group.cells')}>
			<MenuButton
				icon={<Plus size={15} />}
				label={t('cmd.insert')}
				disabled={readOnly}
				items={[
					{ label: t('cmd.insertCells'), onSelect: () => shell.openDialog('insertCells') },
					{ label: t('cmd.insertSheetRows'), onSelect: () => insertCellsCmd(store, 'rows') },
					{ label: t('cmd.insertSheetColumns'), onSelect: () => insertCellsCmd(store, 'cols') },
					{ separator: true },
					{ label: t('cmd.insertSheet'), onSelect: () => insertSheetCmd(store) },
				]}
			/>
			<MenuButton
				icon={<Trash2 size={15} />}
				label={t('cmd.delete')}
				disabled={readOnly}
				items={[
					{ label: t('cmd.deleteCells'), onSelect: () => shell.openDialog('deleteCells') },
					{ label: t('cmd.deleteSheetRows'), onSelect: () => deleteCellsCmd(store, 'rows') },
					{ label: t('cmd.deleteSheetColumns'), onSelect: () => deleteCellsCmd(store, 'cols') },
					{ separator: true },
					{ label: t('cmd.deleteSheet'), danger: true, onSelect: () => deleteSheetCmd(store) },
				]}
			/>
			<MenuButton
				icon={<TableProperties size={15} />}
				label={t('cmd.format')}
				disabled={readOnly}
				items={[
					{ heading: t('cmd.cellSize') },
					{ label: t('cmd.rowHeight'), onSelect: () => shell.openDialog('rowHeight') },
					{ label: t('cmd.autoFitRow'), onSelect: () => autoFitRows(store, active().row) },
					{ label: t('cmd.columnWidth'), onSelect: () => shell.openDialog('columnWidth') },
					{ label: t('cmd.autoFitColumn'), onSelect: () => autoFitColumns(store, active().col) },
					{ heading: t('cmd.visibility') },
					{
						label: t('cmd.hideUnhide'),
						items: [
							{ label: t('cmd.hideRows'), onSelect: () => hideRows(store, true) },
							{ label: t('cmd.hideColumns'), onSelect: () => hideColumns(store, true) },
							{ label: t('cmd.unhideRows'), onSelect: () => hideRows(store, false) },
							{ label: t('cmd.unhideColumns'), onSelect: () => hideColumns(store, false) },
							{ separator: true },
							{
								label: t('cmd.hideSheet'),
								onSelect: () => hideSheetCmd(store, store.state.activeSheet, true),
							},
							{
								label: t('cmd.unhideSheet'),
								disabled: !store.workbook.sheets.some((s) => s.state === 'hidden'),
								onSelect: () => shell.openDialog('unhideSheet'),
							},
						],
					},
					{ heading: t('cmd.organizeSheets') },
					{
						label: t('cmd.renameSheet'),
						icon: <SquarePen size={14} />,
						onSelect: () => shell.openDialog('renameSheet'),
					},
					{ separator: true },
					{
						label: t('cmd.formatCells'),
						shortcut: 'Ctrl+1',
						onSelect: () => shell.openDialog('formatCells'),
					},
				]}
			/>
		</RibbonGroup>
	);
}

export function EditingGroup() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const readOnly = store.readOnly;
	return (
		<RibbonGroup label={t('group.editing')}>
			<SplitButton
				icon={<Sigma size={15} />}
				label={t('cmd.autoSum')}
				disabled={readOnly}
				shortcut='Alt+='
				onClick={() => autoSum(store)}
				items={[
					{ label: t('cmd.sum'), onSelect: () => autoSum(store, 'SUM') },
					{ label: t('cmd.average'), onSelect: () => autoSum(store, 'AVERAGE') },
					{ label: t('cmd.countNumbers'), onSelect: () => autoSum(store, 'COUNT') },
					{ label: t('cmd.max'), onSelect: () => autoSum(store, 'MAX') },
					{ label: t('cmd.min'), onSelect: () => autoSum(store, 'MIN') },
					{ separator: true },
					{ label: t('cmd.moreFunctions'), onSelect: () => shell.openDialog('insertFunction') },
				]}
			/>
			<MenuButton
				icon={<ArrowDownToLine size={15} />}
				label={t('cmd.fill')}
				showLabel={false}
				disabled={readOnly}
				items={[
					{
						label: t('cmd.fillDown'),
						icon: <ArrowDownToLine size={14} />,
						shortcut: 'Ctrl+D',
						onSelect: () => fillDownOrRight(store, 'down'),
					},
					{
						label: t('cmd.fillRight'),
						icon: <ArrowRightToLine size={14} />,
						shortcut: 'Ctrl+R',
						onSelect: () => fillDownOrRight(store, 'right'),
					},
				]}
			/>
			<MenuButton
				icon={<Eraser size={15} />}
				label={t('cmd.clear')}
				showLabel={false}
				disabled={readOnly}
				items={[
					{ label: t('cmd.clearAll'), onSelect: () => clearAll(store, 'all') },
					{ label: t('cmd.clearFormats'), onSelect: () => clearAll(store, 'formats') },
					{ label: t('cmd.clearContents'), onSelect: () => clearAll(store, 'contents') },
					{ label: t('cmd.clearComments'), onSelect: () => clearAll(store, 'comments') },
					{ label: t('cmd.clearHyperlinks'), onSelect: () => clearAll(store, 'hyperlinks') },
				]}
			/>
			<MenuButton
				icon={<ArrowDownAZ size={15} />}
				label={t('cmd.sortFilter')}
				disabled={readOnly}
				items={[
					{
						label: t('cmd.sortAZ'),
						icon: <ArrowDownAZ size={14} />,
						onSelect: () => sortSelection(store, false),
					},
					{
						label: t('cmd.sortZA'),
						icon: <ArrowDownZA size={14} />,
						onSelect: () => sortSelection(store, true),
					},
					{ label: t('cmd.customSort'), onSelect: () => shell.openDialog('sort') },
					{ separator: true },
					{
						label: t('cmd.filter'),
						icon: <Funnel size={14} />,
						shortcut: 'Ctrl+Shift+L',
						onSelect: () => toggleAutoFilter(store),
					},
				]}
			/>
			<MenuButton
				icon={<Search size={15} />}
				label={t('cmd.findSelect')}
				items={[
					{ label: t('cmd.find'), shortcut: 'Ctrl+F', onSelect: () => shell.openDialog('find') },
					{
						label: t('cmd.replace'),
						shortcut: 'Ctrl+H',
						disabled: readOnly,
						onSelect: () => shell.openDialog('replace'),
					},
					{ label: t('cmd.goTo'), shortcut: 'Ctrl+G', onSelect: () => shell.openDialog('goTo') },
				]}
			/>
		</RibbonGroup>
	);
}
