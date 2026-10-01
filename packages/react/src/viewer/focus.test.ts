import { afterEach, describe, expect, it } from 'vitest';

import { focusGridIn, isGridKeyTarget, keepGridFocus } from './focus';

function setup() {
	const root = document.createElement('div');
	root.innerHTML = `
		<div class="ribbon">
			<button id="bold"><svg id="icon"></svg><span id="label">Bold</span></button>
			<input id="font-size" />
			<input id="gridlines" type="checkbox" />
			<input id="zoom" type="range" />
			<div id="empty"></div>
			<div id="tab" draggable="true">Sheet1</div>
		</div>
		<div role="dialog"><button id="dialog-ok">OK</button></div>
		<div data-xv-popover><button id="swatch">Red</button></div>
		<textarea class="xv-cell-editor" id="editor"></textarea>
		<textarea class="xv-focus-sink" id="sink"></textarea>
	`;
	document.body.append(root);
	const $ = (id: string) => root.querySelector<HTMLElement>(`#${id}`)!;
	const press = (id: string, init: MouseEventInit = {}) => {
		const e = new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0, ...init });
		$(id).dispatchEvent(e);
		keepGridFocus(e, root, false);
		return e;
	};
	return { root, $, press };
}

afterEach(() => {
	document.body.innerHTML = '';
});

describe('keepGridFocus', () => {
	it('keeps the grid focused when a toolbar button or empty chrome is pressed', () => {
		const { $, press } = setup();
		$('sink').focus();
		for (const id of ['bold', 'icon', 'label', 'empty', 'gridlines']) {
			expect(press(id).defaultPrevented).toBe(true);
			expect(document.activeElement).toBe($('sink'));
		}
	});

	it('brings focus back to the grid from other chrome controls', () => {
		const { $, press } = setup();
		$('font-size').focus();
		expect(press('bold').defaultPrevented).toBe(true);
		expect(document.activeElement).toBe($('sink'));
		$('sink').blur();
		press('empty');
		expect(document.activeElement).toBe($('sink'));
	});

	it('lets text fields, sliders and drag sources take focus', () => {
		const { $, press } = setup();
		$('sink').focus();
		for (const id of ['font-size', 'zoom', 'tab', 'editor']) {
			expect(press(id).defaultPrevented).toBe(false);
		}
	});

	it('keeps a cell editor focused while editing', () => {
		const { root, $ } = setup();
		$('editor').focus();
		const e = new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 });
		$('bold').dispatchEvent(e);
		keepGridFocus(e, root, true);
		expect(e.defaultPrevented).toBe(true);
		expect(document.activeElement).toBe($('editor'));
	});

	it('never moves focus to the grid while editing', () => {
		const { root, $ } = setup();
		$('font-size').focus();
		const e = new MouseEvent('mousedown', { bubbles: true, cancelable: true, button: 0 });
		$('bold').dispatchEvent(e);
		keepGridFocus(e, root, true);
		expect(e.defaultPrevented).toBe(false);
		expect(document.activeElement).toBe($('font-size'));
	});

	it('leaves dialogs and popovers their own focus', () => {
		const { $, press } = setup();
		$('font-size').focus();
		expect(press('dialog-ok').defaultPrevented).toBe(false);
		expect(press('swatch').defaultPrevented).toBe(false);
		// ...but clicking a popover button while the grid has the keyboard keeps it there.
		$('sink').focus();
		expect(press('swatch').defaultPrevented).toBe(true);
		expect(document.activeElement).toBe($('sink'));
	});

	it('ignores other buttons and presses on a scrollbar', () => {
		const { $, press } = setup();
		$('sink').focus();
		expect(press('bold', { button: 2 }).defaultPrevented).toBe(false);
		const empty = $('empty');
		Object.defineProperty(empty, 'clientWidth', { value: 100 });
		Object.defineProperty(empty, 'clientHeight', { value: 50 });
		expect(press('empty', { clientX: 105, clientY: 10 }).defaultPrevented).toBe(false);
		expect(press('empty', { clientX: 50, clientY: 10 }).defaultPrevented).toBe(true);
	});
});

describe('focus helpers', () => {
	it('focuses the sink of a viewer', () => {
		const { root, $ } = setup();
		focusGridIn(root);
		expect(document.activeElement).toBe($('sink'));
		expect(() => focusGridIn(null)).not.toThrow();
	});

	it('recognizes keys sent to the grid', () => {
		const { $ } = setup();
		expect(isGridKeyTarget($('sink'))).toBe(true);
		expect(isGridKeyTarget($('editor'))).toBe(false);
		expect(isGridKeyTarget(null)).toBe(false);
	});
});
