import { describe, expect, it } from 'vitest';

import { mapGridKey, type KeyInput } from './keymap';

const k = (key: string, mods: Partial<KeyInput> = {}): KeyInput => ({
	key,
	ctrlKey: false,
	metaKey: false,
	shiftKey: false,
	altKey: false,
	...mods,
});

describe('mapGridKey', () => {
	it('maps navigation keys', () => {
		expect(mapGridKey(k('ArrowDown'))).toEqual({
			id: 'move',
			dir: 'down',
			extend: false,
			jump: false,
		});
		expect(mapGridKey(k('ArrowRight', { ctrlKey: true, shiftKey: true }))).toEqual({
			id: 'move',
			dir: 'right',
			extend: true,
			jump: true,
		});
		expect(mapGridKey(k('Tab', { shiftKey: true }))).toEqual({ id: 'advance', dir: 'left' });
		expect(mapGridKey(k('PageDown', { ctrlKey: true }))).toEqual({ id: 'sheet', delta: 1 });
		expect(mapGridKey(k('Home', { metaKey: true }))).toEqual({
			id: 'home',
			ctrl: true,
			extend: false,
		});
	});

	it('maps editing and formatting shortcuts', () => {
		expect(mapGridKey(k('b', { ctrlKey: true }))).toEqual({ id: 'bold' });
		expect(mapGridKey(k('%', { ctrlKey: true, shiftKey: true, code: 'Digit5' }))).toEqual({
			id: 'numberFormat',
			code: '0%',
		});
		expect(mapGridKey(k('=', { altKey: true, code: 'Equal' }))).toEqual({ id: 'autoSum' });
		expect(mapGridKey(k('F2'))).toEqual({ id: 'edit' });
		expect(mapGridKey(k('x'))).toEqual({ id: 'type', text: 'x' });
		expect(mapGridKey(k('Shift'))).toBeNull();
	});
});

describe('outline shortcuts', () => {
	it('maps Alt+Shift+Right / Left to group / ungroup', () => {
		const k = (key: string) => ({
			key,
			ctrlKey: false,
			metaKey: false,
			shiftKey: true,
			altKey: true,
		});
		expect(mapGridKey(k('ArrowRight'))).toEqual({ id: 'group' });
		expect(mapGridKey(k('ArrowLeft'))).toEqual({ id: 'ungroup' });
	});
});
