import { forwardRef, type KeyboardEvent } from 'react';

import { useT } from '../../../i18n';
import { copySelection, pasteData } from '../../commands/clipboard-commands';
import { beginEdit } from '../../commands/edit-commands';
import { useViewerStore } from '../../state/use-store';

interface GridFocusSinkProps {
	onKeyDown: (e: KeyboardEvent<HTMLElement>) => void;
}

/**
 * The hidden textarea that holds the grid's keyboard focus: typing starts an
 * edit, and copy, cut and paste go through the browser's clipboard events.
 */
export const GridFocusSink = forwardRef<HTMLTextAreaElement, GridFocusSinkProps>(
	function GridFocusSink({ onKeyDown }, ref) {
		const t = useT();
		const store = useViewerStore();
		return (
			<textarea
				ref={ref}
				className='xv-focus-sink'
				aria-label={t('grid.spreadsheet')}
				tabIndex={0}
				autoComplete='off'
				spellCheck={false}
				onKeyDown={onKeyDown}
				onInput={(e) => {
					const el = e.currentTarget;
					const text = el.value;
					el.value = '';
					if (text && !store.readOnly && !(e.nativeEvent as InputEvent).isComposing) {
						beginEdit(store, { initialText: text, mode: 'enter' });
					}
				}}
				onCompositionStart={() => {
					if (!store.readOnly) {
						beginEdit(store, { initialText: '', mode: 'enter' });
					}
				}}
				onCopy={(e) => {
					e.preventDefault();
					const data = copySelection(store, false);
					e.clipboardData.setData('text/plain', data.text);
					e.clipboardData.setData('text/html', data.html);
				}}
				onCut={(e) => {
					e.preventDefault();
					if (store.readOnly) {
						return;
					}
					const data = copySelection(store, true);
					e.clipboardData.setData('text/plain', data.text);
					e.clipboardData.setData('text/html', data.html);
				}}
				onPaste={(e) => {
					e.preventDefault();
					pasteData(store, {
						text: e.clipboardData.getData('text/plain') || undefined,
						html: e.clipboardData.getData('text/html') || undefined,
					});
				}}
			/>
		);
	},
);
