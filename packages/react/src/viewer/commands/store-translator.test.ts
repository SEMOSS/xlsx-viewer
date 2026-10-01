import { describe, expect, it } from 'vitest';
import { createWorkbook } from 'xlsx-viewer-core';

import { createTranslator, de, en, fr } from '../../i18n';
import { ViewerStore } from '../state/viewer-store';
import { traceDependents } from './audit-commands';
import { toggleFontFlag } from './format-commands';
import { renameSheetCmd } from './sheet-commands';

describe('the store translator', () => {
	it('defaults to English', () => {
		const store = new ViewerStore(createWorkbook());
		traceDependents(store);
		expect(store.state.message).toBe(en['message.noDependents']);
	});

	it('translates status messages, errors and undo labels of commands', () => {
		const store = new ViewerStore(createWorkbook({ sheetNames: ['A', 'B'] }), {
			t: createTranslator({}, 'de'),
		});
		traceDependents(store);
		expect(store.state.message).toBe(de['message.noDependents']);
		expect(renameSheetCmd(store, 0, 'B')).toBe(de['message.sheetNameTaken']);
		toggleFontFlag(store, 'bold');
		expect(store.getSnapshot().history.undo).toBe('Fett');
	});

	it('follows a translator swapped at runtime, like a locale change in ExcelViewer', () => {
		const store = new ViewerStore(createWorkbook());
		store.t = createTranslator({}, 'fr');
		traceDependents(store);
		expect(store.state.message).toBe(fr['message.noDependents']);
		expect(store.state.message).not.toBe(en['message.noDependents']);
	});
});
