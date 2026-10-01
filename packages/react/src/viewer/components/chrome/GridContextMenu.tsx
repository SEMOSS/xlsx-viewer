import {
	ArrowDownAZ,
	ArrowDownZA,
	ClipboardPaste,
	Copy,
	ExternalLink,
	Link,
	Scissors,
	StickyNote,
	Table2,
	Tag,
	Unlink,
} from 'lucide-react';

import { useT } from '../../../i18n';
import { copyToSystem, pasteFromSystem } from '../../commands/clipboard-commands';
import { clearSelectionContents } from '../../commands/edit-commands';
import {
	followHyperlink,
	hyperlinkForCell,
	removeHyperlinksCmd,
} from '../../commands/hyperlink-commands';
import { hideColumns, hideRows } from '../../commands/layout-commands';
import { deleteCellsCmd, insertCellsCmd, sortSelection } from '../../commands/structure-commands';
import { useViewerStore } from '../../state/use-store';
import { useShell } from '../../shell-context';
import { MenuList, type MenuItem } from '../ui/Menu';
import { Popover } from '../ui/Popover';

export interface ContextMenuState {
	x: number;
	y: number;
	target: 'cell' | 'rowHeader' | 'colHeader' | 'corner';
}

/** Right-click menu of the grid (cells, row headers, column headers). */
export function GridContextMenu({
	state,
	onClose,
}: {
	state: ContextMenuState | null;
	onClose: () => void;
}) {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	if (!state) {
		return null;
	}
	const ro = store.readOnly;
	const sheet = store.activeSheet;
	const { row, col } = store.selectionOf().active;
	const hasNote = sheet.notes.some((n) => n.ref.row === row && n.ref.col === col);
	const link = hyperlinkForCell(store, row, col);
	const clipboard: MenuItem[] = [
		{
			label: t('menu.cut'),
			icon: <Scissors size={14} />,
			shortcut: 'Ctrl+X',
			disabled: ro,
			onSelect: () => void copyToSystem(store, true),
		},
		{
			label: t('menu.copy'),
			icon: <Copy size={14} />,
			shortcut: 'Ctrl+C',
			onSelect: () => void copyToSystem(store, false),
		},
		{
			label: t('menu.paste'),
			icon: <ClipboardPaste size={14} />,
			shortcut: 'Ctrl+V',
			disabled: ro,
			onSelect: () => void pasteFromSystem(store),
		},
		{
			label: t('menu.pasteSpecial'),
			disabled: ro,
			items: [
				{ label: t('cmd.pasteValues'), onSelect: () => void pasteFromSystem(store, 'values') },
				{ label: t('cmd.pasteFormulas'), onSelect: () => void pasteFromSystem(store, 'formulas') },
				{ label: t('cmd.pasteFormatting'), onSelect: () => void pasteFromSystem(store, 'formats') },
				{
					label: t('cmd.pasteTranspose'),
					onSelect: () => void pasteFromSystem(store, 'transpose'),
				},
			],
		},
		{ separator: true },
	];
	let items: MenuItem[];
	if (state.target === 'rowHeader' || state.target === 'colHeader') {
		const rows = state.target === 'rowHeader';
		items = [
			...clipboard,
			{
				label: t('menu.insertRows'),
				disabled: ro,
				onSelect: () => insertCellsCmd(store, rows ? 'rows' : 'cols'),
			},
			{
				label: t('menu.deleteRows'),
				disabled: ro,
				onSelect: () => deleteCellsCmd(store, rows ? 'rows' : 'cols'),
			},
			{
				label: t('menu.clearContents'),
				disabled: ro,
				onSelect: () => clearSelectionContents(store),
			},
			{ separator: true },
			{
				label: t('menu.formatCells'),
				disabled: ro,
				onSelect: () => shell.openDialog('formatCells'),
			},
			{
				label: rows ? t('menu.rowHeight') : t('menu.columnWidth'),
				disabled: ro,
				onSelect: () => shell.openDialog(rows ? 'rowHeight' : 'columnWidth'),
			},
			{
				label: t('menu.hide'),
				disabled: ro,
				onSelect: () => (rows ? hideRows(store, true) : hideColumns(store, true)),
			},
			{
				label: t('menu.unhide'),
				disabled: ro,
				onSelect: () => (rows ? hideRows(store, false) : hideColumns(store, false)),
			},
		];
	} else {
		items = [
			...clipboard,
			{
				label: t('menu.insert'),
				disabled: ro,
				items: [
					{ label: t('cmd.insertSheetRows'), onSelect: () => insertCellsCmd(store, 'rows') },
					{ label: t('cmd.insertSheetColumns'), onSelect: () => insertCellsCmd(store, 'cols') },
				],
			},
			{
				label: t('menu.delete'),
				disabled: ro,
				items: [
					{ label: t('cmd.deleteSheetRows'), onSelect: () => deleteCellsCmd(store, 'rows') },
					{ label: t('cmd.deleteSheetColumns'), onSelect: () => deleteCellsCmd(store, 'cols') },
				],
			},
			{
				label: t('menu.clearContents'),
				disabled: ro,
				onSelect: () => clearSelectionContents(store),
			},
			{ separator: true },
			{
				label: t('menu.sort'),
				disabled: ro,
				items: [
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
				],
			},
			{ separator: true },
			{
				label: hasNote ? t('menu.editNote') : t('menu.newNote'),
				icon: <StickyNote size={14} />,
				disabled: ro,
				onSelect: () => shell.openDialog('note'),
			},
			{
				label: t('menu.formatCells'),
				icon: <Table2 size={14} />,
				shortcut: 'Ctrl+1',
				disabled: ro,
				onSelect: () => shell.openDialog('formatCells'),
			},
			{
				label: t('menu.defineName'),
				icon: <Tag size={14} />,
				disabled: ro,
				onSelect: () => shell.openDialog('defineName'),
			},
			...(link
				? [
						{
							label: t('hyperlink.edit'),
							icon: <Link size={14} />,
							shortcut: 'Ctrl+K',
							disabled: ro,
							onSelect: () => shell.openDialog('hyperlink'),
						},
						{
							label: t('hyperlink.open'),
							icon: <ExternalLink size={14} />,
							onSelect: () => void followHyperlink(store, link),
						},
						{
							label: t('hyperlink.remove'),
							icon: <Unlink size={14} />,
							disabled: ro,
							onSelect: () => removeHyperlinksCmd(store),
						},
					]
				: [
						{
							label: t('menu.link'),
							icon: <Link size={14} />,
							shortcut: 'Ctrl+K',
							disabled: ro,
							onSelect: () => shell.openDialog('hyperlink'),
						},
					]),
		];
	}
	return (
		<Popover
			open
			placement='point'
			point={{ x: state.x, y: state.y }}
			onClose={onClose}
			role='menu'
		>
			<MenuList items={items} onClose={onClose} />
		</Popover>
	);
}
