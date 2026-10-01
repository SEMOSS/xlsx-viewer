import type { FilterTarget } from 'xlsx-viewer-core';
import { filterColumnValues, getFilter, setFilterColumn } from 'xlsx-viewer-core';
import { ArrowDownAZ, ArrowDownZA, Search } from 'lucide-react';
import { useMemo, useState } from 'react';

import { useT } from '../../../i18n';
import { useViewerStore } from '../../state/use-store';
import { sortSelection } from '../../commands/structure-commands';

/** Data values are never empty, so '' stands for the (Blanks) entry of the checklist. */
const BLANKS = '';

interface FilterMenuProps {
	target: FilterTarget;
	colId: number;
	onClose: () => void;
}

/** The AutoFilter dropdown: sort, search and a value checklist. */
export function FilterMenu({ target, colId, onClose }: FilterMenuProps) {
	const t = useT();
	const store = useViewerStore();
	const sheetIndex = store.state.activeSheet;
	const sheet = store.activeSheet;
	const filter = getFilter(sheet, target);
	const data = useMemo(
		() =>
			filter
				? filterColumnValues(store.workbook, sheet, filter.ref, colId)
				: { values: [], hasBlank: false },
		[store, sheet, filter, colId],
	);
	const existing = filter?.columns.find((c) => c.colId === colId)?.filters;
	const [checked, setChecked] = useState<Set<string>>(() => {
		if (!existing) {
			return new Set([
				...data.values.map((v) => v.toLowerCase()),
				...(data.hasBlank ? ['(blank)'] : []),
			]);
		}
		return new Set([
			...existing.values.map((v) => v.toLowerCase()),
			...(existing.blank ? ['(blank)'] : []),
		]);
	});
	const [query, setQuery] = useState('');
	if (!filter) {
		return null;
	}
	const items = [...data.values, ...(data.hasBlank ? [BLANKS] : [])];
	const keyOf = (v: string) => (v === BLANKS ? '(blank)' : v.toLowerCase());
	const labelOf = (v: string) => (v === BLANKS ? t('filter.blanks') : v);
	const visible = items.filter((v) => labelOf(v).toLowerCase().includes(query.toLowerCase()));
	const allChecked = items.every((v) => checked.has(keyOf(v)));

	const sort = (descending: boolean) => {
		const col = filter.ref.left + colId;
		store.setSelection({
			active: { row: filter.ref.top + 1, col },
			anchor: { row: filter.ref.top + 1, col },
			ranges: [filter.ref],
		});
		sortSelection(store, descending);
		onClose();
	};

	const apply = () => {
		const everything = allChecked;
		store.transact(store.t('cmd.filter'), (tx) => {
			setFilterColumn(
				tx,
				store.workbook,
				sheetIndex,
				target,
				colId,
				everything
					? null
					: {
							filters: {
								values: data.values.filter((v) => checked.has(v.toLowerCase())),
								blank: checked.has('(blank)') || undefined,
							},
						},
			);
		});
		onClose();
	};

	return (
		<div
			className='flex w-[240px] flex-col gap-1 p-1.5 text-[13px]'
			role='dialog'
			aria-label={t('cmd.filter')}
		>
			<button
				type='button'
				className='flex items-center gap-2 rounded-sm px-2 py-1.5 text-left hover:bg-accent'
				onClick={() => sort(false)}
			>
				<ArrowDownAZ size={14} /> {t('cmd.sortAZ')}
			</button>
			<button
				type='button'
				className='flex items-center gap-2 rounded-sm px-2 py-1.5 text-left hover:bg-accent'
				onClick={() => sort(true)}
			>
				<ArrowDownZA size={14} /> {t('cmd.sortZA')}
			</button>
			{existing && (
				<button
					type='button'
					className='rounded-sm px-2 py-1.5 text-left hover:bg-accent'
					onClick={() => {
						store.transact(store.t('undo.clearFilter'), (tx) =>
							setFilterColumn(tx, store.workbook, sheetIndex, target, colId, null),
						);
						onClose();
					}}
				>
					{t('undo.clearFilter')}
				</button>
			)}
			<div className='my-1 h-px bg-border' />
			<label className='flex items-center gap-1.5 rounded-sm border border-input px-1.5'>
				<Search size={13} className='text-muted-foreground' />
				<input
					className='h-7 min-w-0 flex-1 bg-transparent outline-none'
					placeholder={t('filter.search')}
					value={query}
					onChange={(e) => setQuery(e.target.value)}
					aria-label={t('filter.searchValues')}
				/>
			</label>
			<div className='max-h-[220px] overflow-y-auto rounded-sm border border-border py-1'>
				<label className='flex items-center gap-2 px-2 py-0.5'>
					<input
						type='checkbox'
						checked={allChecked}
						onChange={(e) => setChecked(e.target.checked ? new Set(items.map(keyOf)) : new Set())}
					/>
					{t('filter.selectAll')}
				</label>
				{visible.map((v) => (
					<label key={v} className='flex items-center gap-2 px-2 py-0.5'>
						<input
							type='checkbox'
							checked={checked.has(keyOf(v))}
							onChange={(e) => {
								const next = new Set(checked);
								if (e.target.checked) {
									next.add(keyOf(v));
								} else {
									next.delete(keyOf(v));
								}
								setChecked(next);
							}}
						/>
						<span className='truncate'>{labelOf(v)}</span>
					</label>
				))}
			</div>
			<div className='flex justify-end gap-2 pt-1'>
				<button
					type='button'
					className='rounded-md bg-primary px-3 py-1 text-primary-foreground disabled:opacity-50'
					disabled={checked.size === 0}
					onClick={apply}
				>
					{t('dialog.ok')}
				</button>
				<button
					type='button'
					className='rounded-md border border-border px-3 py-1 hover:bg-accent'
					onClick={onClose}
				>
					{t('dialog.cancel')}
				</button>
			</div>
		</div>
	);
}
