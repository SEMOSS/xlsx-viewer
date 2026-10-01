/**
 * Keyboard focus in the viewer. As in Excel, the grid owns the keyboard between
 * edits: its focus sink (a hidden textarea) keeps focus when toolbar buttons,
 * sheet tabs or empty chrome are clicked, so typing goes straight into the
 * active cell. Text fields, sliders and drag sources take focus as usual, and
 * dialogs, popovers and the backstage manage their own.
 */

/** The grid's focus sink. */
export const FOCUS_SINK = '.xv-focus-sink';

/** Controls that need real focus when clicked. */
const TAKES_FOCUS = [
	'input:not([type="checkbox"]):not([type="radio"]):not([type="button"]):not([type="submit"]):not([type="reset"])',
	'textarea',
	'select',
	'[contenteditable]:not([contenteditable="false"])',
	'[draggable="true"]',
].join(', ');

/** Elements that hold the keyboard for the sheet: the focus sink and the cell editors. */
const KEYBOARD_OWNER = '.xv-focus-sink, .xv-cell-editor, .xv-formula-input';

/** Layers above the sheet that keep their own focus. */
const OVERLAY = '[role="dialog"], [data-xv-modal], [data-xv-popover], .xv-backstage';

/** A press on an element's own scrollbar (beyond its client area). */
function onScrollbar(e: MouseEvent, target: Element): boolean {
	return (
		target instanceof HTMLElement &&
		target.clientWidth > 0 &&
		(e.offsetX > target.clientWidth || e.offsetY > target.clientHeight)
	);
}

/** Focus the grid of a viewer. */
export function focusGridIn(root: HTMLElement | null | undefined): void {
	root?.querySelector<HTMLElement>(FOCUS_SINK)?.focus({ preventScroll: true });
}

/**
 * Mouse-down handler for the viewer root. When the grid (or a cell editor)
 * has the keyboard, clicks on buttons and empty chrome leave it there; when
 * focus is elsewhere (a toolbar field, nowhere), they bring it back to the
 * grid. Pass `editing` while a cell edit is in progress: focus then is never
 * moved to the grid.
 */
export function keepGridFocus(e: MouseEvent, root: HTMLElement, editing: boolean): void {
	const target = e.target;
	if (e.button !== 0 || e.defaultPrevented || !(target instanceof Element)) {
		return;
	}
	if (!root.contains(target) || target.closest(TAKES_FOCUS) || onScrollbar(e, target)) {
		return;
	}
	const sink = root.querySelector<HTMLElement>(FOCUS_SINK);
	if (!sink) {
		return;
	}
	const active = document.activeElement;
	if (active instanceof Element && root.contains(active) && active.closest(KEYBOARD_OWNER)) {
		e.preventDefault();
		return;
	}
	if (editing || target.closest(OVERLAY)) {
		return;
	}
	e.preventDefault();
	sink.focus({ preventScroll: true });
}

/** Whether a keyboard event was sent to a grid's focus sink. */
export function isGridKeyTarget(target: EventTarget | null): boolean {
	return target instanceof Element && target.matches(FOCUS_SINK);
}
