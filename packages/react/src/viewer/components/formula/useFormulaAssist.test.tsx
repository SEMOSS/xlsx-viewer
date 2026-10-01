import { act, renderHook } from '@testing-library/react';
import type { KeyboardEvent, ReactNode } from 'react';
import { createWorkbook } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { beginEdit, updateEditText } from '../../commands/edit-commands';
import { ViewerStoreContext, useStoreSelector } from '../../state/use-store';
import { ViewerStore } from '../../state/viewer-store';
import { useFormulaAssist } from './useFormulaAssist';

function setup(text: string, source: 'cell' | 'formulaBar' = 'cell') {
	const wb = createWorkbook();
	wb.definedNames.push(
		{ name: 'Kiwi_Sales', formula: 'Sheet1!$B$2:$B$9' },
		{ name: 'Kilos', formula: '0.07' },
	);
	const store = new ViewerStore(wb);
	beginEdit(store, { initialText: text, source });
	const wrapper = ({ children }: { children: ReactNode }) => (
		<ViewerStoreContext.Provider value={store}>{children}</ViewerStoreContext.Provider>
	);
	const hook = renderHook(
		() => {
			const edit = useStoreSelector((s) => s.ui.edit);
			return useFormulaAssist(store, edit, edit?.source === 'cell');
		},
		{ wrapper },
	);
	return { store, hook };
}

const key = (k: string) =>
	({
		key: k,
		altKey: false,
		ctrlKey: false,
		metaKey: false,
		preventDefault() {},
	}) as unknown as KeyboardEvent<HTMLElement>;

describe('useFormulaAssist', () => {
	it('completes defined names and inserts the chosen one', () => {
		const { store, hook } = setup('=SUM(Kiw');
		expect(hook.result.current.completions.map((c) => c.name)).toEqual(['Kiwi_Sales']);
		act(() => {
			hook.result.current.onKeyDown(key('Tab'));
		});
		expect(store.state.edit?.text).toBe('=SUM(Kiwi_Sales');
		expect(store.state.edit?.caret).toBe('=SUM(Kiwi_Sales'.length);
	});

	it('moves the highlight with the arrow keys and wraps', () => {
		const { store, hook } = setup('=Ki');
		expect(hook.result.current.completions.map((c) => c.name)).toEqual(['Kiwi_Sales', 'Kilos']);
		act(() => {
			hook.result.current.onKeyDown(key('ArrowDown'));
		});
		expect(hook.result.current.highlighted).toBe(1);
		act(() => {
			hook.result.current.onKeyDown(key('ArrowDown'));
		});
		expect(hook.result.current.highlighted).toBe(0);
		act(() => {
			hook.result.current.onKeyDown(key('ArrowUp'));
		});
		act(() => {
			hook.result.current.onKeyDown(key('Tab'));
		});
		expect(store.state.edit?.text).toBe('=Kilos');
	});

	it('closes on Escape until the text changes', () => {
		const { store, hook } = setup('=Ki');
		act(() => {
			expect(hook.result.current.onKeyDown(key('Escape'))).toBe(true);
		});
		expect(hook.result.current.completions).toHaveLength(0);
		// A second Escape is not consumed (it cancels the edit).
		expect(hook.result.current.onKeyDown(key('Escape'))).toBe(false);
		act(() => updateEditText(store, '=Kiw', 4));
		expect(hook.result.current.completions).toHaveLength(1);
	});

	it('lists functions with an opening parenthesis on accept and shows argument help', () => {
		const { store, hook } = setup('=SUMI');
		expect(hook.result.current.completions.map((c) => c.name)).toEqual(['SUMIF', 'SUMIFS']);
		act(() => {
			hook.result.current.onKeyDown(key('Tab'));
		});
		expect(store.state.edit?.text).toBe('=SUMIF(');
		expect(hook.result.current.hint?.name).toBe('SUMIF');
		expect(hook.result.current.hint?.parts[0]).toMatchObject({ label: 'range', current: true });
	});

	it('cycles reference anchoring with F4', () => {
		const { store, hook } = setup('=A1+B2');
		act(() => {
			hook.result.current.onKeyDown(key('F4'));
		});
		expect(store.state.edit?.text).toBe('=A1+$B$2');
	});

	it('stays quiet for plain text and for the other editor', () => {
		expect(setup('Ki').hook.result.current.completions).toHaveLength(0);
		const other = setup('=Ki', 'formulaBar');
		expect(other.hook.result.current.completions).toHaveLength(0);
		expect(other.hook.result.current.onKeyDown(key('F4'))).toBe(false);
	});
});
