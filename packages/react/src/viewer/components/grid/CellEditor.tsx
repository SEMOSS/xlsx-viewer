import { useEffect, useLayoutEffect, useRef, type CSSProperties, type KeyboardEvent } from 'react';

import { useT } from '../../../i18n';
import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import type { MergeIndex } from '../../../engine/layout/style-at';
import { styleIdAt } from '../../../engine/layout/style-at';
import { fillToCss } from '../../../engine/render/fills';
import { fontToCss } from '../../../engine/render/fonts';
import { measureLines } from '../../../engine/render/text-measure';
import { cancelEdit, commitEdit, typeInCell, updateEditText } from '../../commands/edit-commands';
import { pointMove } from '../../commands/point-commands';
import type { EditState } from '../../state/types';
import { useViewerStore } from '../../state/use-store';
import { FormulaAssistPopup } from '../formula/FormulaAssistPopup';
import { useFormulaAssist } from '../formula/useFormulaAssist';
import { resolveColorCss } from 'xlsx-viewer-core';

interface CellEditorProps {
	edit: EditState;
	geometry: SheetGeometry;
	merges: MergeIndex;
	/** Right edge of the visible area in sheet coordinates (the editor stops growing there). */
	maxRight: number;
	onDone: () => void;
}

const ARROW_DIRS = {
	ArrowUp: 'up',
	ArrowDown: 'down',
	ArrowLeft: 'left',
	ArrowRight: 'right',
} as const;

/** The in-cell editor: a textarea over the cell that grows right, then wraps. */
export function CellEditor({ edit, geometry: g, merges, maxRight, onDone }: CellEditorProps) {
	const t = useT();
	const store = useViewerStore();
	const ref = useRef<HTMLTextAreaElement>(null);
	const assist = useFormulaAssist(store, edit, edit.source === 'cell');
	const sheet = store.workbook.sheets[edit.sheet];
	const merge = merges.find(edit.row, edit.col);
	const rect = merge ? g.rangeRect(merge) : g.cellRect(edit.row, edit.col);
	const style = store.styles.resolve(styleIdAt(sheet, edit.row, edit.col));
	const font = fontToCss(style.font, g.zoom);
	const colorCtx = {
		theme: store.workbook.theme,
		indexedColors: store.workbook.styles.indexedColors,
	};
	const color = resolveColorCss(style.font.color, colorCtx, '#000000');
	// Excel's editor shows the cell's own fill, so light text on a dark fill stays readable.
	const fill = fillToCss(style.fill, colorCtx);
	const fontString = `${font.fontStyle} ${font.fontWeight} ${font.fontSize}px ${font.fontFamily}`;
	const textWidth =
		measureLines(
			edit.text + (edit.suggestion ? edit.suggestion.slice(edit.text.length) : '') || ' ',
			fontString,
		) + 12;
	const width = Math.max(
		rect.width + 1,
		Math.min(textWidth, Math.max(rect.width, maxRight - rect.x)),
	);
	const lineCount = edit.text.split('\n').length;
	const lineHeight = Math.round(font.fontSize * 1.25);
	const wrapsNeeded = textWidth > width ? Math.ceil(textWidth / Math.max(1, width)) : 1;
	const height = Math.max(rect.height + 1, (lineCount + wrapsNeeded - 1) * lineHeight + 4);

	// AutoComplete: the rest of the suggested entry follows the caret, selected.
	const completion =
		edit.suggestion && edit.source === 'cell' ? edit.suggestion.slice(edit.text.length) : '';
	const shown = edit.text + completion;

	useLayoutEffect(() => {
		const el = ref.current;
		if (!el) {
			return;
		}
		if (edit.source === 'cell' && document.activeElement !== el) {
			el.focus({ preventScroll: true });
		}
		const end = edit.caret + (completion ? completion.length : 0);
		if (el.selectionStart !== edit.caret || el.selectionEnd !== end) {
			el.setSelectionRange(edit.caret, end);
		}
	}, [edit.caret, edit.source, completion]);

	useEffect(() => {
		if (edit.source === 'cell') {
			ref.current?.focus({ preventScroll: true });
		}
	}, [edit.source]);

	const onKeyDown = (e: KeyboardEvent<HTMLTextAreaElement>) => {
		if (e.nativeEvent.isComposing || assist.onKeyDown(e)) {
			return;
		}
		const finish = (ok: boolean) => {
			if (ok) {
				onDone();
			}
		};
		if (e.key === 'Enter') {
			if (e.altKey) {
				e.preventDefault();
				const el = e.currentTarget;
				const start = el.selectionStart;
				const text = edit.text.slice(0, start) + '\n' + edit.text.slice(el.selectionEnd);
				updateEditText(store, text, start + 1);
				return;
			}
			e.preventDefault();
			finish(
				commitEdit(store, {
					move: e.shiftKey ? 'up' : 'down',
					fillSelection: e.ctrlKey || e.metaKey,
				}),
			);
			return;
		}
		if (e.key === 'Tab') {
			e.preventDefault();
			finish(commitEdit(store, { move: e.shiftKey ? 'left' : 'right' }));
			return;
		}
		if (e.key === 'Escape') {
			e.preventDefault();
			cancelEdit(store);
			onDone();
			return;
		}
		if (e.key === 'F2') {
			e.preventDefault();
			store.setUi({ edit: { ...edit, mode: edit.mode === 'enter' ? 'edit' : 'enter' } });
			return;
		}
		const dir = ARROW_DIRS[e.key as keyof typeof ARROW_DIRS];
		if (dir && edit.mode === 'enter') {
			e.preventDefault();
			if (pointMove(store, dir, e.shiftKey)) {
				return;
			}
			finish(commitEdit(store, { move: dir }));
		}
	};

	const css: CSSProperties = {
		position: 'absolute',
		left: rect.x - 1,
		top: rect.y - 1,
		width,
		height,
		fontFamily: font.fontFamily,
		fontSize: `${font.fontSize}px`,
		fontWeight: font.fontWeight,
		fontStyle: font.fontStyle,
		color,
		backgroundColor: fill?.backgroundColor ?? '#ffffff',
		backgroundImage: fill?.backgroundImage,
		lineHeight: `${lineHeight}px`,
		padding: `${Math.max(0, rect.height - lineHeight - 3)}px 2px 1px 2px`,
		zIndex: 20,
	};
	if (lineCount > 1 || wrapsNeeded > 1) {
		css.paddingTop = 1;
	}

	return (
		<>
			<textarea
				ref={ref}
				className='xv-cell-editor'
				style={css}
				value={shown}
				spellCheck={false}
				autoComplete='off'
				aria-label={t('grid.cellEditor')}
				data-xv-editor
				onChange={(e) => typeInCell(store, e.target.value, e.target.selectionStart)}
				onSelect={(e) => {
					const el = e.currentTarget;
					if (
						el.selectionStart === el.selectionEnd &&
						(el.selectionStart !== edit.caret || edit.suggestion)
					) {
						// Moving the caret drops an AutoComplete suggestion.
						store.setUi({
							edit: { ...edit, caret: el.selectionStart, point: undefined, suggestion: undefined },
						});
					}
				}}
				onKeyDown={onKeyDown}
				onPointerDown={(e) => {
					e.stopPropagation();
					if (edit.mode === 'enter') {
						store.setUi({ edit: { ...edit, mode: 'edit' } });
					}
				}}
			/>
			<FormulaAssistPopup assist={assist} anchorRef={ref} />
		</>
	);
}
