import type { FunctionCategory, FunctionInfo } from 'xlsx-viewer-core';
import {
	formatFunctionSyntax,
	FUNCTION_CATEGORIES,
	getFunctionInfo,
	listFunctions,
} from 'xlsx-viewer-core';
import { useId, useMemo, useState } from 'react';

import { FUNCTION_CATEGORY_KEYS, useT } from '../../../i18n';
import {
	insertFunctionCall,
	recentFunctions,
	rememberFunction,
	searchFunctionCatalog,
} from '../../commands/insert-function-commands';
import { useViewerStore } from '../../state/use-store';
import { DialogButton } from '../ui/Dialog';
import { FIELD, FormDialog, ListBox } from './dialog-ui';
import type { DialogProps } from './types';

type CategoryChoice = 'recommended' | 'recent' | 'all' | FunctionCategory;

/**
 * Formulas > Insert Function (fx): find a function by search or category
 * and start it in the active cell.
 */
export function InsertFunctionDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const searchId = useId();
	const all = useMemo(() => listFunctions(), []);
	const [query, setQuery] = useState('');
	const [results, setResults] = useState<FunctionInfo[] | null>(null);
	const [category, setCategory] = useState<CategoryChoice>('recent');
	const listFor = (c: CategoryChoice): FunctionInfo[] => {
		if (c === 'recommended') {
			return results ?? [];
		}
		if (c === 'recent') {
			return recentFunctions()
				.map((name) => getFunctionInfo(name))
				.filter((f): f is FunctionInfo => f !== undefined);
		}
		return c === 'all' ? all : listFunctions(c);
	};
	const items = listFor(category);
	const [selected, setSelected] = useState<string | null>(() => items[0]?.name ?? null);
	const info = selected ? getFunctionInfo(selected) : undefined;

	const go = () => {
		const found = searchFunctionCatalog(query, all);
		setResults(found);
		setCategory('recommended');
		setSelected(found[0]?.name ?? null);
	};

	const run = (name: string | null) => {
		if (!name || store.readOnly) {
			return;
		}
		rememberFunction(name);
		onClose();
		insertFunctionCall(store, name);
	};

	return (
		<FormDialog
			title={t('cmd.insertFunction')}
			onClose={onClose}
			onOk={() => run(selected)}
			okDisabled={!selected || store.readOnly}
			width={500}
		>
			<div className='flex flex-col gap-1'>
				<label htmlFor={searchId}>{t('dialog.insertFunction.search')}</label>
				<div className='flex gap-2' data-enter='self'>
					<input
						id={searchId}
						className={FIELD}
						value={query}
						placeholder={t('dialog.insertFunction.searchPlaceholder')}
						onChange={(e) => setQuery(e.target.value)}
						onKeyDown={(e) => {
							if (e.key === 'Enter') {
								e.preventDefault();
								go();
							}
						}}
					/>
					<DialogButton onClick={go} disabled={query.trim() === ''}>
						{t('dialog.insertFunction.go')}
					</DialogButton>
				</div>
			</div>
			<label className='flex items-center gap-2'>
				<span className='whitespace-nowrap'>{t('dialog.insertFunction.category')}</span>
				<select
					className={`${FIELD} w-[220px]`}
					value={category}
					onChange={(e) => {
						const next = e.target.value as CategoryChoice;
						setCategory(next);
						setSelected(listFor(next)[0]?.name ?? null);
					}}
				>
					{results && <option value='recommended'>{t('dialog.insertFunction.recommended')}</option>}
					<option value='recent'>{t('dialog.insertFunction.mostRecentlyUsed')}</option>
					<option value='all'>{t('dialog.insertFunction.all')}</option>
					{FUNCTION_CATEGORIES.map((c) => (
						<option key={c} value={c}>
							{t(FUNCTION_CATEGORY_KEYS[c])}
						</option>
					))}
				</select>
			</label>
			<div className='flex flex-col gap-1'>
				<span>{t('dialog.insertFunction.selectLabel')}</span>
				<ListBox
					label={t('dialog.insertFunction.select')}
					items={items}
					getKey={(f) => f.name}
					getText={(f) => f.name}
					value={selected}
					onChange={setSelected}
					onActivate={run}
					render={(f) => f.name}
					className='h-[170px]'
					empty={
						<div className='px-2 py-1 text-muted-foreground'>
							{t(
								category === 'recommended'
									? 'dialog.insertFunction.noMatch'
									: 'dialog.insertFunction.noFunctions',
							)}
						</div>
					}
				/>
			</div>
			<div className='min-h-[64px]' aria-live='polite'>
				{info ? (
					<>
						<p className='font-semibold' data-testid='function-syntax'>
							{formatFunctionSyntax(info)}
						</p>
						<p className='mt-1 text-[12.5px] text-muted-foreground'>{info.description}</p>
					</>
				) : (
					selected && <p className='text-muted-foreground'>{t('dialog.insertFunction.noHelp')}</p>
				)}
			</div>
		</FormDialog>
	);
}
