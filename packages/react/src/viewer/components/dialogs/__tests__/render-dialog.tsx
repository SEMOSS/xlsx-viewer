import { render, type RenderResult } from '@testing-library/react';
import type { ComponentType } from 'react';
import { vi, type Mock } from 'vitest';
import type { Workbook } from 'xlsx-viewer-core';
import { createWorkbook, EditTransaction, parseRange, setCellInput } from 'xlsx-viewer-core';

import { ShellContext, type ShellActions } from '../../../shell-context';
import { ViewerStoreContext } from '../../../state/use-store';
import { ViewerStore } from '../../../state/viewer-store';
import type { DialogProps } from '../types';

/**
 * Test helpers for dialog smoke tests: render a dialog component inside a
 * store and shell context, seed cells and select ranges.
 */

export interface RenderDialogOptions {
	workbook?: Workbook;
	store?: ViewerStore;
	payload?: unknown;
	readOnly?: boolean;
	shell?: Partial<ShellActions>;
}

export function makeShell(overrides: Partial<ShellActions> = {}): ShellActions {
	return {
		openDialog: vi.fn(),
		save: vi.fn(),
		download: vi.fn(),
		print: vi.fn(),
		notify: vi.fn(),
		themeMode: 'light',
		...overrides,
	};
}

export interface RenderedDialog extends RenderResult {
	store: ViewerStore;
	onClose: Mock<() => void>;
	shell: ShellActions;
}

export function renderDialog(
	Component: ComponentType<DialogProps>,
	opts: RenderDialogOptions = {},
): RenderedDialog {
	const store =
		opts.store ?? new ViewerStore(opts.workbook ?? createWorkbook(), { readOnly: opts.readOnly });
	const onClose = vi.fn<() => void>();
	const shell = makeShell(opts.shell);
	const utils = render(
		<ViewerStoreContext.Provider value={store}>
			<ShellContext.Provider value={shell}>
				<Component onClose={onClose} payload={opts.payload} author='Tester' />
			</ShellContext.Provider>
		</ViewerStoreContext.Provider>,
	);
	return { ...utils, store, onClose, shell };
}

/** Type cell entries (`{ A1: 'Name', B2: '5', C3: '=A1' }`) on a sheet as if typed. */
export function typeCells(
	workbook: Workbook,
	entries: Record<string, string | number>,
	sheetIndex = 0,
): void {
	const tx = new EditTransaction(workbook, 'seed');
	for (const [ref, value] of Object.entries(entries)) {
		const r = parseRange(ref);
		if (!r) {
			throw new Error(`bad ref ${ref}`);
		}
		setCellInput(tx, workbook, sheetIndex, r.top, r.left, String(value));
	}
	tx.commit();
}

/** Select a range (`'A1:B3'`) on the active sheet, with the active cell at its top-left. */
export function selectRange(store: ViewerStore, ref: string): void {
	const range = parseRange(ref);
	if (!range) {
		throw new Error(`bad ref ${ref}`);
	}
	const active = { row: range.top, col: range.left };
	store.setSelection({ active, anchor: active, ranges: [range] });
}
