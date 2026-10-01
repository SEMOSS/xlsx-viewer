import { cellKey, formatRange } from 'xlsx-viewer-core';
import { Check, ChevronDown, ChevronUp, X } from 'lucide-react';
import { useEffect, useRef, useState, type KeyboardEvent } from 'react';

import { useT } from '../../../i18n';
import { currentRange } from '../../../engine/selection/selection';
import {
	beginEdit,
	cancelEdit,
	commitEdit,
	editTextForCell,
	updateEditText,
} from '../../commands/edit-commands';
import { goToReference } from '../../commands/selection-commands';
import { styleIdAt } from '../../../engine/layout/style-at';
import { measureLines } from '../../../engine/render/text-measure';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { useReturnToGrid, useShell } from '../../shell-context';
import { FormulaAssistPopup } from '../formula/FormulaAssistPopup';
import { useFormulaAssist } from '../formula/useFormulaAssist';

function NameBox() {
	const t = useT();
	const store = useViewerStore();
	const returnToGrid = useReturnToGrid();
	const label = useStoreSelector((snap, st) => {
		const sel = st.selectionOf(snap.ui.activeSheet);
		const r = currentRange(sel);
		const name = st.workbook.definedNames.find((n) => {
			const ref = n.formula.replace(/^.*!/, '').replace(/\$/g, '');
			return !n.hidden && !n.builtin && ref === formatRange(r);
		});
		if (name && (r.top !== r.bottom || r.left !== r.right)) {
			return name.name;
		}
		return formatRange({
			top: sel.active.row,
			left: sel.active.col,
			bottom: sel.active.row,
			right: sel.active.col,
		});
	});
	const [text, setText] = useState(label);
	const [focused, setFocused] = useState(false);
	useEffect(() => {
		if (!focused) {
			setText(label);
		}
	}, [label, focused]);
	const onKey = (e: KeyboardEvent<HTMLInputElement>) => {
		if (e.key === 'Enter') {
			e.preventDefault();
			if (!goToReference(store, text)) {
				store.setUi({ message: t('message.invalidReferenceAt', { ref: text }) });
				setText(label);
			}
			returnToGrid(e.currentTarget);
		} else if (e.key === 'Escape') {
			setText(label);
			returnToGrid(e.currentTarget);
		}
	};
	return (
		<div className='relative flex h-[24px] w-[104px] shrink-0 items-center rounded-sm border border-input bg-card focus-within:border-primary'>
			<input
				aria-label={t('formula.nameBox')}
				title={t('formula.nameBox')}
				value={text}
				className='h-full min-w-0 flex-1 bg-transparent px-1.5 text-[12px] text-foreground outline-none'
				onFocus={(e) => {
					setFocused(true);
					e.currentTarget.select();
				}}
				onBlur={() => setFocused(false)}
				onChange={(e) => setText(e.target.value)}
				onKeyDown={onKey}
			/>
			<ChevronDown size={13} className='mr-1 text-muted-foreground' aria-hidden />
		</div>
	);
}

/** Distance from the input's left edge to `index` in `text` (to place the autocomplete list). */
function caretOffset(el: HTMLTextAreaElement | null, text: string, index: number): number {
	if (!el) {
		return 0;
	}
	const lineStart = text.lastIndexOf('\n', index - 1) + 1;
	const cs = getComputedStyle(el);
	const width = measureLines(text.slice(lineStart, index), `${cs.fontSize} ${cs.fontFamily}`);
	return Math.max(
		0,
		Math.min(el.clientWidth - 40, parseFloat(cs.paddingLeft) + width - el.scrollLeft),
	);
}

/** Name box, edit buttons and the formula input. */
export function FormulaBar() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const inputRef = useRef<HTMLTextAreaElement>(null);
	const edit = useStoreSelector((s) => s.ui.edit);
	const expanded = useStoreSelector((s) => s.ui.formulaBarExpanded);
	const display = useStoreSelector((snap, st) => {
		void snap.docVersion;
		const sheet = st.workbook.sheets[snap.ui.activeSheet];
		const { row, col } = st.selectionOf(snap.ui.activeSheet).active;
		const cell = sheet.cells.get(cellKey(row, col));
		const style = st.styles.resolve(styleIdAt(sheet, row, col));
		if (style.protection.hidden && sheet.protection?.sheet) {
			return '';
		}
		return editTextForCell(cell, style, st.workbook.date1904);
	});
	const value = edit ? edit.text : display;
	const editingHere = edit?.source === 'formulaBar';
	const assist = useFormulaAssist(store, edit, editingHere);
	const assistOffset =
		edit && editingHere ? caretOffset(inputRef.current, edit.text, assist.start) : 0;

	useEffect(() => {
		const el = inputRef.current;
		if (el && editingHere && document.activeElement === el && el.selectionStart !== edit.caret) {
			el.setSelectionRange(edit.caret, edit.caret);
		}
	}, [editingHere, edit?.caret]);

	const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
		if (!store.state.edit || e.nativeEvent.isComposing || assist.onKeyDown(e)) {
			return;
		}
		if (e.key === 'Enter' && !e.altKey) {
			e.preventDefault();
			commitEdit(store, {
				move: e.shiftKey ? 'up' : 'down',
				fillSelection: e.ctrlKey || e.metaKey,
			});
			inputRef.current?.blur();
		} else if (e.key === 'Enter' && e.altKey) {
			e.preventDefault();
			const el = e.currentTarget;
			const s = el.selectionStart;
			updateEditText(
				store,
				`${store.state.edit.text.slice(0, s)}\n${store.state.edit.text.slice(el.selectionEnd)}`,
				s + 1,
			);
		} else if (e.key === 'Escape') {
			e.preventDefault();
			cancelEdit(store);
			inputRef.current?.blur();
		} else if (e.key === 'Tab') {
			e.preventDefault();
			commitEdit(store, { move: e.shiftKey ? 'left' : 'right' });
			inputRef.current?.blur();
		}
	};

	return (
		<div className='xv-formula-bar flex shrink-0 items-start gap-1 border-b border-border bg-card px-2 py-[3px]'>
			<NameBox />
			<div className='flex h-[24px] items-center gap-0.5 px-1 text-muted-foreground'>
				<button
					type='button'
					aria-label={t('formula.cancel')}
					title={t('formula.cancel')}
					disabled={!edit}
					className='rounded-sm p-0.5 hover:bg-accent disabled:opacity-30'
					onClick={() => cancelEdit(store)}
				>
					<X size={15} />
				</button>
				<button
					type='button'
					aria-label={t('formula.enter')}
					title={t('formula.enter')}
					disabled={!edit}
					className='rounded-sm p-0.5 hover:bg-accent disabled:opacity-30'
					onClick={() => commitEdit(store, { move: 'none' })}
				>
					<Check size={15} />
				</button>
				<button
					type='button'
					aria-label={t('formula.insertFunction')}
					title={t('formula.insertFunction')}
					disabled={store.readOnly}
					className='rounded-sm px-1 py-0.5 font-serif text-[13px] italic hover:bg-accent disabled:opacity-30'
					onClick={() => shell.openDialog('insertFunction')}
				>
					fx
				</button>
			</div>
			<textarea
				ref={inputRef}
				aria-label={t('formulaBar.label')}
				value={value}
				readOnly={store.readOnly}
				spellCheck={false}
				rows={1}
				className={`xv-formula-input min-w-0 flex-1 resize-none rounded-sm border border-input bg-card px-1.5 py-[3px] font-[Consolas,Menlo,monospace] text-[12.5px] leading-[16px] text-foreground outline-none focus:border-primary ${
					expanded ? 'h-[72px]' : 'h-[24px] overflow-hidden whitespace-pre'
				}`}
				onFocus={() => {
					if (!store.readOnly && !store.state.edit) {
						beginEdit(store, { source: 'formulaBar', mode: 'edit' });
					} else if (store.state.edit && store.state.edit.source !== 'formulaBar') {
						store.setUi({ edit: { ...store.state.edit, source: 'formulaBar', mode: 'edit' } });
					}
				}}
				onChange={(e) => updateEditText(store, e.target.value, e.target.selectionStart)}
				onSelect={(e) => {
					const el = e.currentTarget;
					const current = store.state.edit;
					if (
						current &&
						editingHere &&
						el.selectionStart === el.selectionEnd &&
						el.selectionStart !== current.caret
					) {
						store.setUi({ edit: { ...current, caret: el.selectionStart, point: undefined } });
					}
				}}
				onKeyDown={onKeyDown}
			/>
			{editingHere && (
				<FormulaAssistPopup assist={assist} anchorRef={inputRef} offsetX={assistOffset} />
			)}
			<button
				type='button'
				aria-label={t('formula.expand')}
				title={t('formula.expand')}
				className='mt-[3px] rounded-sm p-0.5 text-muted-foreground hover:bg-accent'
				onClick={() => store.setUi({ formulaBarExpanded: !expanded })}
			>
				{expanded ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
			</button>
		</div>
	);
}
