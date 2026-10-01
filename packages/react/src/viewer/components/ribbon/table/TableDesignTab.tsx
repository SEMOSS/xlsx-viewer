import type { Table } from 'xlsx-viewer-core';
import { getTableStyle } from 'xlsx-viewer-core';
import { ArrowLeftRight, CopyMinus } from 'lucide-react';
import { useEffect, useMemo, useState, type KeyboardEvent } from 'react';

import { useT, type TranslationKey } from '../../../../i18n';
import { buildTableThumbnail } from '../../../../engine/render/table-thumbnail';
import {
	activeTable,
	convertActiveTableToRange,
	renameActiveTable,
	setActiveTableStyle,
	setTableOption,
	tableOptionValue,
	type TableOption,
} from '../../../commands/table-commands';
import { useReturnToGrid, useShell } from '../../../shell-context';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import { isDialogAvailable } from '../../dialogs/DialogHost';
import { TableStyleGallery, tableStyleLabel } from '../../tables/TableStyleGallery';
import { TableThumbnailView } from '../../tables/TableThumbnailView';
import { MenuButton, RibbonButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';

const OPTIONS: Array<{ option: TableOption; key: TranslationKey }> = [
	{ option: 'headerRow', key: 'table.headerRow' },
	{ option: 'totalRow', key: 'table.totalRow' },
	{ option: 'showRowStripes', key: 'table.bandedRows' },
	{ option: 'filterButton', key: 'table.filterButton' },
	{ option: 'showFirstColumn', key: 'table.firstColumn' },
	{ option: 'showLastColumn', key: 'table.lastColumn' },
	{ option: 'showColumnStripes', key: 'table.bandedColumns' },
];

function TableNameBox({ table }: { table: Table }) {
	const t = useT();
	const store = useViewerStore();
	const returnToGrid = useReturnToGrid();
	const [text, setText] = useState(table.name);
	useEffect(() => setText(table.name), [table.name]);
	const commit = () => {
		const error = text.trim() === table.name ? null : renameActiveTable(store, text);
		if (error) {
			store.setUi({ message: error });
			setText(table.name);
		}
	};
	const onKey = (e: KeyboardEvent<HTMLInputElement>) => {
		if (e.key === 'Enter') {
			e.preventDefault();
			commit();
			returnToGrid(e.currentTarget);
		} else if (e.key === 'Escape') {
			setText(table.name);
			returnToGrid(e.currentTarget);
		}
	};
	return (
		<label className='flex flex-col text-[10.5px] leading-[12px] text-muted-foreground'>
			{t('table.name')}
			<input
				value={text}
				disabled={store.readOnly}
				aria-label={t('table.name')}
				className='mt-0.5 h-[19px] w-[112px] rounded-sm border border-input bg-card px-1 text-[12px] text-foreground outline-none focus:border-primary'
				onChange={(e) => setText(e.target.value)}
				onBlur={commit}
				onKeyDown={onKey}
			/>
		</label>
	);
}

/** Contextual ribbon tab shown while the active cell is inside a table. */
export function TableDesignTab() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	// Re-render on any document or selection change; the table object is replaced on edits.
	const current = useStoreSelector(
		(snap, st) => {
			void snap.docVersion;
			void snap.ui;
			return activeTable(st);
		},
		(a, b) => a?.table === b?.table && a?.index === b?.index,
	);
	const theme = store.workbook.theme;
	const styles = store.workbook.styles;
	const styleName = current?.table.style?.name;
	const thumb = useMemo(() => {
		const def = styleName ? getTableStyle(styleName, styles) : undefined;
		return def
			? buildTableThumbnail(
					def,
					current?.table.style ?? {},
					{ theme, indexedColors: styles.indexedColors },
					{ width: 7, height: 4.5 },
				)
			: null;
	}, [styleName, styles, theme, current?.table.style]);

	if (!current) {
		return null;
	}
	const { table } = current;
	const ro = store.readOnly;
	return (
		<>
			<RibbonGroup label={t('group.tableProperties')}>
				<TableNameBox table={table} />
			</RibbonGroup>
			<RibbonGroup label={t('group.tableTools')}>
				<RibbonButton
					icon={<ArrowLeftRight size={15} />}
					label={t('table.convertToRange')}
					showLabel
					disabled={ro}
					onClick={() => {
						if (window.confirm(t('table.convertConfirm'))) {
							convertActiveTableToRange(store);
						}
					}}
				/>
				<RibbonButton
					icon={<CopyMinus size={15} />}
					label={t('table.removeDuplicates')}
					showLabel
					disabled={ro || !isDialogAvailable('removeDuplicates')}
					onClick={() => shell.openDialog('removeDuplicates')}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.tableStyleOptions')}>
				<div className='grid grid-flow-col grid-rows-2 gap-x-3 gap-y-0.5'>
					{OPTIONS.map(({ option, key }) => (
						<label
							key={option}
							className='flex items-center gap-1 text-[11.5px] whitespace-nowrap text-foreground'
						>
							<input
								type='checkbox'
								className='h-3 w-3 accent-[var(--xlsx-selection-border,#107c41)]'
								checked={tableOptionValue(table, option)}
								disabled={ro || (option === 'filterButton' && table.headerRowCount === 0)}
								onChange={(e) => setTableOption(store, option, e.target.checked)}
							/>
							{t(key)}
						</label>
					))}
				</div>
			</RibbonGroup>
			<RibbonGroup label={t('group.tableStyles')}>
				<MenuButton
					icon={thumb ? <TableThumbnailView thumb={thumb} /> : undefined}
					label={styleName ? tableStyleLabel(styleName, t) : t('table.none')}
					disabled={ro}
					renderPanel={(close) => (
						<TableStyleGallery
							current={styleName}
							options={table.style}
							allowNone
							onPick={(name) => {
								setActiveTableStyle(store, name);
								close();
							}}
						/>
					)}
				/>
			</RibbonGroup>
		</>
	);
}
