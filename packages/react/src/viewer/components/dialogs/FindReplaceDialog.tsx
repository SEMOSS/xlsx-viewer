import { a1, setCellInput } from 'xlsx-viewer-core';
import { useMemo, useRef, useState } from 'react';

import { formatNumber, useLocale, useT } from '../../../i18n';
import { findAll, replaceInText, type FindHit, type FindOptions } from '../../../engine/find/find';
import { selectCell } from '../../commands/selection-commands';
import { useViewerStore } from '../../state/use-store';
import { Dialog, DialogButton } from '../ui/Dialog';

const field =
	'w-full rounded-md border border-input bg-card px-2 py-1.5 text-[13px] text-foreground outline-none focus:border-primary';

/** Find and Replace (Ctrl+F / Ctrl+H). */
export function FindReplaceDialog({
	mode,
	onClose,
}: {
	mode: 'find' | 'replace';
	onClose: () => void;
}) {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const [tab, setTab] = useState(mode);
	const [query, setQuery] = useState('');
	const [replacement, setReplacement] = useState('');
	const [matchCase, setMatchCase] = useState(false);
	const [wholeCell, setWholeCell] = useState(false);
	const [within, setWithin] = useState<'sheet' | 'workbook'>('sheet');
	const [lookIn, setLookIn] = useState<'formulas' | 'values'>('formulas');
	const [results, setResults] = useState<FindHit[] | null>(null);
	const cursor = useRef(-1);
	const opts: FindOptions = useMemo(
		() => ({
			query,
			matchCase,
			wholeCell,
			within,
			lookIn: tab === 'replace' ? 'formulas' : lookIn,
		}),
		[query, matchCase, wholeCell, within, lookIn, tab],
	);

	const goTo = (hit: FindHit) => {
		if (hit.sheet !== store.state.activeSheet) {
			store.setUi({ activeSheet: hit.sheet });
		}
		selectCell(store, hit.row, hit.col);
	};

	const findNext = () => {
		const hits = findAll(store.workbook, store.state.activeSheet, opts);
		if (hits.length === 0) {
			store.setUi({ message: t('dialog.find.notFound') });
			return;
		}
		const { row, col } = store.selectionOf().active;
		const sheet = store.state.activeSheet;
		let idx = hits.findIndex(
			(h) =>
				h.sheet > sheet || (h.sheet === sheet && (h.row > row || (h.row === row && h.col > col))),
		);
		if (idx < 0) {
			idx = 0;
		}
		cursor.current = idx;
		goTo(hits[idx]);
	};

	const replaceAll = () => {
		const hits = findAll(store.workbook, store.state.activeSheet, { ...opts, lookIn: 'formulas' });
		if (hits.length === 0) {
			store.setUi({ message: t('dialog.find.nothingToReplace') });
			return;
		}
		store.transact(store.t('undo.replace'), (tx) => {
			for (const h of hits) {
				setCellInput(
					tx,
					store.workbook,
					h.sheet,
					h.row,
					h.col,
					replaceInText(h.text, opts, replacement),
				);
			}
		});
		const count = formatNumber(hits.length, locale);
		store.setUi({
			message: t(hits.length === 1 ? 'dialog.find.replacedOne' : 'dialog.find.replacedOther', {
				count,
			}),
		});
	};

	const replaceOne = () => {
		const { row, col } = store.selectionOf().active;
		const sheet = store.state.activeSheet;
		const hits = findAll(store.workbook, sheet, { ...opts, within: 'sheet', lookIn: 'formulas' });
		const here = hits.find((h) => h.row === row && h.col === col);
		if (here) {
			store.transact(store.t('undo.replace'), (tx) => {
				setCellInput(
					tx,
					store.workbook,
					sheet,
					row,
					col,
					replaceInText(here.text, opts, replacement),
				);
			});
		}
		findNext();
	};

	const tabBtn = (id: 'find' | 'replace', label: string) => (
		<button
			type='button'
			role='tab'
			aria-selected={tab === id}
			className={`px-3 py-1.5 text-[13px] ${tab === id ? 'border-b-2 border-primary font-semibold' : 'text-muted-foreground'}`}
			onClick={() => setTab(id)}
		>
			{label}
		</button>
	);

	return (
		<Dialog
			title={t('dialog.find.title')}
			onClose={onClose}
			width={520}
			footer={
				<>
					{tab === 'replace' && (
						<DialogButton onClick={replaceAll} disabled={!query || store.readOnly}>
							{t('dialog.find.replaceAll')}
						</DialogButton>
					)}
					{tab === 'replace' && (
						<DialogButton onClick={replaceOne} disabled={!query || store.readOnly}>
							{t('dialog.find.replace')}
						</DialogButton>
					)}
					<DialogButton
						onClick={() => setResults(findAll(store.workbook, store.state.activeSheet, opts))}
						disabled={!query}
					>
						{t('dialog.find.findAll')}
					</DialogButton>
					<DialogButton primary onClick={findNext} disabled={!query}>
						{t('dialog.find.findNext')}
					</DialogButton>
					<DialogButton onClick={onClose}>{t('dialog.close')}</DialogButton>
				</>
			}
		>
			<div role='tablist' className='mb-3 flex gap-1 border-b border-border'>
				{tabBtn('find', t('dialog.find.findTab'))}
				{tabBtn('replace', t('dialog.find.replaceTab'))}
			</div>
			<div className='grid grid-cols-[110px_1fr] items-center gap-2'>
				<span>{t('dialog.find.findWhatLabel')}</span>
				<input
					className={field}
					value={query}
					onChange={(e) => setQuery(e.target.value)}
					onKeyDown={(e) => e.key === 'Enter' && findNext()}
					aria-label={t('dialog.find.findWhat')}
				/>
				{tab === 'replace' && (
					<>
						<span>{t('dialog.find.replaceWithLabel')}</span>
						<input
							className={field}
							value={replacement}
							onChange={(e) => setReplacement(e.target.value)}
							aria-label={t('dialog.find.replaceWith')}
						/>
					</>
				)}
				<span>{t('dialog.find.within')}</span>
				<select
					className={field}
					value={within}
					onChange={(e) => setWithin(e.target.value as 'sheet' | 'workbook')}
				>
					<option value='sheet'>{t('dialog.find.withinSheet')}</option>
					<option value='workbook'>{t('dialog.find.withinWorkbook')}</option>
				</select>
				{tab === 'find' && (
					<>
						<span>{t('dialog.find.lookIn')}</span>
						<select
							className={field}
							value={lookIn}
							onChange={(e) => setLookIn(e.target.value as 'formulas' | 'values')}
						>
							<option value='formulas'>{t('dialog.find.lookInFormulas')}</option>
							<option value='values'>{t('dialog.find.lookInValues')}</option>
						</select>
					</>
				)}
			</div>
			<div className='mt-3 flex gap-5'>
				<label className='flex items-center gap-1.5'>
					<input
						type='checkbox'
						checked={matchCase}
						onChange={(e) => setMatchCase(e.target.checked)}
					/>{' '}
					{t('dialog.find.matchCase')}
				</label>
				<label className='flex items-center gap-1.5'>
					<input
						type='checkbox'
						checked={wholeCell}
						onChange={(e) => setWholeCell(e.target.checked)}
					/>{' '}
					{t('dialog.find.matchEntireCell')}
				</label>
			</div>
			{results && (
				<div className='mt-3 max-h-[200px] overflow-y-auto rounded-md border border-border'>
					<div className='sticky top-0 grid grid-cols-[110px_70px_1fr] bg-muted px-2 py-1 text-[12px] font-semibold'>
						<span>{t('dialog.find.columnSheet')}</span>
						<span>{t('dialog.find.columnCell')}</span>
						<span>{t('dialog.find.columnValue')}</span>
					</div>
					{results.map((h) => (
						<button
							key={`${h.sheet}:${h.row}:${h.col}`}
							type='button'
							className='grid w-full grid-cols-[110px_70px_1fr] px-2 py-1 text-left text-[12.5px] hover:bg-accent'
							onClick={() => goTo(h)}
						>
							<span className='truncate'>{store.workbook.sheets[h.sheet].name}</span>
							<span>{a1(h.row, h.col)}</span>
							<span className='truncate'>{h.text}</span>
						</button>
					))}
					<div className='px-2 py-1 text-[12px] text-muted-foreground'>
						{t('dialog.find.cellsFound', { count: formatNumber(results.length, locale) })}
					</div>
				</div>
			)}
		</Dialog>
	);
}
