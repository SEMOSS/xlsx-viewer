import type { GridCommand } from '../../engine/keymap';
import type { ViewerStore } from '../state/viewer-store';

/** Shortcuts without a dedicated implementation reveal the ribbon tab where the command lives. */
export function showRibbonTabFor(store: ViewerStore, action: GridCommand['id']): void {
	switch (action) {
		case 'table':
			store.setUi({ ribbonTab: 'insert' });
			break;
		case 'dropdown':
		case 'repeat':
		default:
			break;
	}
}
