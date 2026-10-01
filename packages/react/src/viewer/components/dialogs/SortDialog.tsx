import { useEffect, useMemo, useRef, useState } from 'react';

import { useT } from '../../../i18n';
import {
	applySortLevels,
	MAX_SORT_LEVELS,
	selectSortBlock,
	sortDialogSetup,
	sortKeyOptions,
	type SortDialogSetup,
	type SortLevel,
} from '../../commands/sort-dialog-commands';
import { useViewerStore } from '../../state/use-store';
import { DialogMessage, FormDialog } from './dialog-ui';
import { SortLevelGrid, type SortLevelRow } from './sort/SortLevelGrid';
import { SortOptionsDialog, type SortOptionsValue } from './sort/SortOptionsDialog';
import { SortToolbar } from './sort/SortToolbar';
import type { DialogProps } from './types';

/** The active cell's column (row when sorting left to right), kept inside the block. */
function defaultKey(setup: SortDialogSetup, byColumns: boolean): number | null {
	const b = setup.block;
	if (!b) {
		return null;
	}
	return byColumns
		? Math.min(Math.max(setup.active.row, b.top), b.bottom)
		: Math.min(Math.max(setup.active.col, b.left), b.right);
}

/**
 * Data > Sort (Custom Sort): sort the selection, or the current region of a
 * single cell, by up to 64 levels.
 */
export function SortDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const [setup] = useState(() => sortDialogSetup(store));
	const nextId = useRef(1);
	const makeLevel = (key: number | null, descending = false): SortLevelRow => ({
		id: nextId.current++,
		key,
		descending,
	});
	const [levels, setLevels] = useState<SortLevelRow[]>(() => [makeLevel(defaultKey(setup, false))]);
	const [selected, setSelected] = useState(0);
	const [hasHeader, setHasHeader] = useState(setup.hasHeader);
	const [options, setOptions] = useState<SortOptionsValue>({
		caseSensitive: false,
		byColumns: false,
	});
	const [optionsOpen, setOptionsOpen] = useState(false);
	const [error, setError] = useState<string | null>(null);
	const { byColumns } = options;
	const headers = hasHeader && !byColumns;
	const block = setup.block;
	// `t` too: the "Column B" / "Row 3" labels follow the UI language.
	const keyOptions = useMemo(
		() => (block ? sortKeyOptions(store, block, { hasHeader: headers, byColumns }) : []),
		[store, block, headers, byColumns, t],
	);

	useEffect(() => {
		if (setup.error || !setup.block) {
			store.setUi({ message: setup.error ?? null });
			onClose();
			return;
		}
		selectSortBlock(store, setup.block, { hasHeader: setup.hasHeader, byColumns: false });
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []);

	if (!block) {
		return null;
	}

	const edit = (next: SortLevelRow[], select: number) => {
		setLevels(next);
		setSelected(Math.max(0, Math.min(select, next.length - 1)));
		setError(null);
	};
	const insertAfter = (level: SortLevelRow) => {
		const at = levels.length ? selected + 1 : 0;
		edit([...levels.slice(0, at), level, ...levels.slice(at)], at);
	};
	const move = (delta: -1 | 1) => {
		const to = selected + delta;
		const next = [...levels];
		[next[selected], next[to]] = [next[to], next[selected]];
		edit(next, to);
	};
	const change = (index: number, patch: Partial<SortLevel>) => {
		edit(
			levels.map((level, i) => (i === index ? { ...level, ...patch } : level)),
			index,
		);
	};
	const setHeaders = (value: boolean) => {
		setHasHeader(value);
		setError(null);
		selectSortBlock(store, block, { hasHeader: value, byColumns });
	};
	const applyOptions = (value: SortOptionsValue) => {
		setOptionsOpen(false);
		setOptions(value);
		if (value.byColumns !== byColumns) {
			// A new orientation sorts by rows instead of columns: start over.
			edit([makeLevel(defaultKey(setup, value.byColumns))], 0);
			selectSortBlock(store, block, {
				hasHeader: hasHeader && !value.byColumns,
				byColumns: value.byColumns,
			});
		}
	};

	const ok = () => {
		if (store.readOnly) {
			onClose();
			return;
		}
		const message = applySortLevels(
			store,
			block,
			levels,
			{ hasHeader: headers, byColumns, caseSensitive: options.caseSensitive },
			keyOptions,
		);
		if (message) {
			setError(message);
			return;
		}
		onClose();
	};

	const canAdd = levels.length < MAX_SORT_LEVELS;
	return (
		<>
			<FormDialog
				title={t('dialog.sort.title')}
				onClose={onClose}
				onOk={ok}
				okDisabled={store.readOnly}
				width={640}
				initialFocusSelector='[role="grid"] select'
			>
				<SortToolbar
					canAdd={canAdd}
					hasLevel={levels.length > 0}
					canMoveUp={selected > 0}
					canMoveDown={selected < levels.length - 1}
					onAdd={() => insertAfter(makeLevel(null))}
					onDelete={() =>
						edit(
							levels.filter((_, i) => i !== selected),
							selected,
						)
					}
					onCopy={() => {
						const source = levels[selected];
						insertAfter(makeLevel(source.key, source.descending));
					}}
					onMove={move}
					onOptions={() => setOptionsOpen(true)}
					hasHeader={headers}
					headerDisabled={byColumns}
					onHeaderChange={setHeaders}
				/>
				<SortLevelGrid
					levels={levels}
					selected={selected}
					onSelect={setSelected}
					onChange={change}
					keyOptions={keyOptions}
					byColumns={byColumns}
				/>
				{error && <DialogMessage>{error}</DialogMessage>}
			</FormDialog>
			{optionsOpen && (
				<SortOptionsDialog
					caseSensitive={options.caseSensitive}
					byColumns={byColumns}
					onOk={applyOptions}
					onClose={() => setOptionsOpen(false)}
				/>
			)}
		</>
	);
}
