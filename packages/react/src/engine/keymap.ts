/**
 * Keyboard shortcuts of the grid (when not editing), as a pure decision
 * function: key event fields -> command id. Mirrors Excel for Windows,
 * with Cmd treated like Ctrl on macOS.
 */
export type GridCommand =
	| { id: 'move'; dir: 'up' | 'down' | 'left' | 'right'; extend: boolean; jump: boolean }
	| { id: 'advance'; dir: 'up' | 'down' | 'left' | 'right' }
	| { id: 'page'; dir: 'up' | 'down' | 'left' | 'right'; extend: boolean }
	| { id: 'home'; ctrl: boolean; extend: boolean }
	| { id: 'end'; extend: boolean }
	| { id: 'sheet'; delta: 1 | -1 }
	| { id: 'edit' }
	| { id: 'editClear' }
	| { id: 'clear' }
	| { id: 'escape' }
	| { id: 'selectAll' }
	| { id: 'selectColumn' }
	| { id: 'selectRow' }
	| { id: 'copy' }
	| { id: 'cut' }
	| { id: 'paste' }
	| { id: 'undo' }
	| { id: 'redo' }
	| { id: 'repeat' }
	| { id: 'bold' }
	| { id: 'italic' }
	| { id: 'underline' }
	| { id: 'strike' }
	| { id: 'formatCells' }
	| { id: 'find' }
	| { id: 'replace' }
	| { id: 'goTo' }
	| { id: 'fillDown' }
	| { id: 'fillRight' }
	| { id: 'insertDate' }
	| { id: 'insertTime' }
	| { id: 'numberFormat'; code: string }
	| { id: 'showFormulas' }
	| { id: 'save' }
	| { id: 'open' }
	| { id: 'print' }
	| { id: 'hyperlink' }
	| { id: 'note' }
	| { id: 'calculate' }
	| { id: 'filter' }
	| { id: 'autoSum' }
	| { id: 'deleteCells' }
	| { id: 'insertCells' }
	| { id: 'newSheet' }
	| { id: 'hideRows' }
	| { id: 'hideCols' }
	| { id: 'unhideRows' }
	| { id: 'unhideCols' }
	| { id: 'table' }
	| { id: 'dropdown' }
	| { id: 'contextMenu' }
	| { id: 'group' }
	| { id: 'ungroup' }
	| { id: 'type'; text: string };

export interface KeyInput {
	key: string;
	code?: string;
	ctrlKey: boolean;
	metaKey: boolean;
	shiftKey: boolean;
	altKey: boolean;
}

const ARROWS: Record<string, 'up' | 'down' | 'left' | 'right'> = {
	ArrowUp: 'up',
	ArrowDown: 'down',
	ArrowLeft: 'left',
	ArrowRight: 'right',
};

export function mapGridKey(e: KeyInput): GridCommand | null {
	const ctrl = e.ctrlKey || e.metaKey;
	const { shiftKey: shift, altKey: alt } = e;
	const key = e.key;
	const lower = key.length === 1 ? key.toLowerCase() : key;

	if (ARROWS[key]) {
		if (alt && key === 'ArrowDown') {
			return { id: 'dropdown' };
		}
		if (alt && shift && !ctrl && (key === 'ArrowRight' || key === 'ArrowLeft')) {
			return { id: key === 'ArrowRight' ? 'group' : 'ungroup' };
		}
		return { id: 'move', dir: ARROWS[key], extend: shift, jump: ctrl };
	}
	switch (key) {
		case 'Enter':
			return ctrl ? null : { id: 'advance', dir: shift ? 'up' : 'down' };
		case 'Tab':
			return { id: 'advance', dir: shift ? 'left' : 'right' };
		case 'PageDown':
			return ctrl
				? { id: 'sheet', delta: 1 }
				: { id: 'page', dir: alt ? 'right' : 'down', extend: shift };
		case 'PageUp':
			return ctrl
				? { id: 'sheet', delta: -1 }
				: { id: 'page', dir: alt ? 'left' : 'up', extend: shift };
		case 'Home':
			return { id: 'home', ctrl, extend: shift };
		case 'End':
			return ctrl ? { id: 'end', extend: shift } : null;
		case 'F2':
			return shift ? { id: 'note' } : { id: 'edit' };
		case 'F4':
			return ctrl ? null : { id: 'repeat' };
		case 'F5':
			return { id: 'goTo' };
		case 'F9':
			return { id: 'calculate' };
		case 'F10':
			return shift ? { id: 'contextMenu' } : null;
		case 'F11':
			return shift ? { id: 'newSheet' } : null;
		case 'ContextMenu':
			return { id: 'contextMenu' };
		case 'Delete':
			return { id: 'clear' };
		case 'Backspace':
			return { id: 'editClear' };
		case 'Escape':
			return { id: 'escape' };
		case ' ':
			if (ctrl && !shift) {
				return { id: 'selectColumn' };
			}
			if (shift && !ctrl) {
				return { id: 'selectRow' };
			}
			if (ctrl && shift) {
				return { id: 'selectAll' };
			}
			break;
		default:
			break;
	}
	if (alt && !ctrl && (key === '=' || e.code === 'Equal')) {
		return { id: 'autoSum' };
	}
	if (ctrl) {
		if (shift) {
			switch (e.code) {
				case 'Digit4':
					return { id: 'numberFormat', code: '"$"#,##0.00_);[Red]\\("$"#,##0.00\\)' };
				case 'Digit5':
					return { id: 'numberFormat', code: '0%' };
				case 'Digit3':
					return { id: 'numberFormat', code: 'd-mmm-yy' };
				case 'Digit2':
					return { id: 'numberFormat', code: 'h:mm AM/PM' };
				case 'Digit1':
					return { id: 'numberFormat', code: '#,##0.00' };
				case 'Digit6':
					return { id: 'numberFormat', code: '0.00E+00' };
				case 'Backquote':
					return { id: 'numberFormat', code: 'General' };
				case 'Semicolon':
					return { id: 'insertTime' };
				case 'Equal':
					return { id: 'insertCells' };
				case 'Digit9':
					return { id: 'unhideRows' };
				case 'Digit0':
					return { id: 'unhideCols' };
				default:
					break;
			}
			if (lower === 'l') {
				return { id: 'filter' };
			}
			if (lower === 'z') {
				return { id: 'redo' };
			}
		}
		switch (lower) {
			case 'a':
				return { id: 'selectAll' };
			case 'c':
				return { id: 'copy' };
			case 'x':
				return { id: 'cut' };
			case 'v':
				return { id: 'paste' };
			case 'z':
				return { id: 'undo' };
			case 'y':
				return { id: 'redo' };
			case 'b':
			case '2':
				return { id: 'bold' };
			case 'i':
			case '3':
				return { id: 'italic' };
			case 'u':
			case '4':
				return { id: 'underline' };
			case '5':
				return { id: 'strike' };
			case '1':
				return { id: 'formatCells' };
			case 'f':
				return { id: 'find' };
			case 'h':
				return { id: 'replace' };
			case 'g':
				return { id: 'goTo' };
			case 'd':
				return { id: 'fillDown' };
			case 'r':
				return { id: 'fillRight' };
			case ';':
				return { id: 'insertDate' };
			case '`':
				return { id: 'showFormulas' };
			case 's':
				return { id: 'save' };
			case 'o':
				return { id: 'open' };
			case 'p':
				return { id: 'print' };
			case 'k':
				return { id: 'hyperlink' };
			case 't':
			case 'l':
				return { id: 'table' };
			case '-':
				return { id: 'deleteCells' };
			case '9':
				return { id: 'hideRows' };
			case '0':
				return { id: 'hideCols' };
			default:
				return null;
		}
	}
	if (!alt && key.length === 1) {
		return { id: 'type', text: key };
	}
	return null;
}
