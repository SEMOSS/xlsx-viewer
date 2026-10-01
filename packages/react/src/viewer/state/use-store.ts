import { createContext, useCallback, useContext, useRef, useSyncExternalStore } from 'react';

import type { StoreSnapshot } from './types';
import type { ViewerStore } from './viewer-store';

export const ViewerStoreContext = createContext<ViewerStore | null>(null);

/** The store of the enclosing `<ExcelViewer>`. */
export function useViewerStore(): ViewerStore {
	const store = useContext(ViewerStoreContext);
	if (!store) {
		throw new Error('useViewerStore must be used inside <ExcelViewer>');
	}
	return store;
}

/**
 * Subscribe to a slice of the store snapshot. The component re-renders only
 * when `isEqual(previous, next)` is false (default: Object.is).
 */
export function useStoreSelector<T>(
	selector: (snapshot: StoreSnapshot, store: ViewerStore) => T,
	isEqual: (a: T, b: T) => boolean = Object.is,
): T {
	const store = useViewerStore();
	const cache = useRef<{ snapshot: StoreSnapshot; value: T } | null>(null);
	const getValue = useCallback(() => {
		const snapshot = store.getSnapshot();
		const prev = cache.current;
		if (prev && prev.snapshot === snapshot) {
			return prev.value;
		}
		const value = selector(snapshot, store);
		if (prev && isEqual(prev.value, value)) {
			cache.current = { snapshot, value: prev.value };
			return prev.value;
		}
		cache.current = { snapshot, value };
		return value;
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [store, selector, isEqual]);
	return useSyncExternalStore(store.subscribe, getValue, getValue);
}

/** Shallow equality for objects / arrays of primitives. */
export function shallowEqual<T>(a: T, b: T): boolean {
	if (Object.is(a, b)) {
		return true;
	}
	if (typeof a !== 'object' || typeof b !== 'object' || a === null || b === null) {
		return false;
	}
	const ka = Object.keys(a as object);
	const kb = Object.keys(b as object);
	if (ka.length !== kb.length) {
		return false;
	}
	for (const k of ka) {
		if (!Object.is((a as Record<string, unknown>)[k], (b as Record<string, unknown>)[k])) {
			return false;
		}
	}
	return true;
}
