import { useId, useState } from 'react';

import { useT, type TranslationKey } from '../../../i18n';
import { currentRange } from '../../../engine/selection/selection';
import {
	deleteCellsShiftCmd,
	insertCellsShiftCmd,
	wholeAxisOfRanges,
} from '../../commands/cell-shift-commands';
import { deleteCellsCmd, insertCellsCmd } from '../../commands/structure-commands';
import type { ViewerStore } from '../../state/viewer-store';
import { useViewerStore } from '../../state/use-store';
import { FormDialog, GroupBox, Radio } from './dialog-ui';
import type { DialogProps } from './types';

/**
 * Home > Insert > Insert Cells... and Home > Delete > Delete Cells...
 * (Ctrl+Shift+= / Ctrl+-): choose how the surrounding cells move, or insert
 * / delete entire rows or columns instead.
 */

type InsertChoice = 'right' | 'down' | 'rows' | 'cols';
type DeleteChoice = 'left' | 'up' | 'rows' | 'cols';

interface Choice<T extends string> {
	value: T;
	labelKey: TranslationKey;
}

const INSERT_CHOICES: readonly Choice<InsertChoice>[] = [
	{ value: 'right', labelKey: 'dialog.insertCells.shiftRight' },
	{ value: 'down', labelKey: 'dialog.insertCells.shiftDown' },
	{ value: 'rows', labelKey: 'dialog.insertCells.entireRow' },
	{ value: 'cols', labelKey: 'dialog.insertCells.entireColumn' },
];

const DELETE_CHOICES: readonly Choice<DeleteChoice>[] = [
	{ value: 'left', labelKey: 'dialog.deleteCells.shiftLeft' },
	{ value: 'up', labelKey: 'dialog.deleteCells.shiftUp' },
	{ value: 'rows', labelKey: 'dialog.deleteCells.entireRow' },
	{ value: 'cols', labelKey: 'dialog.deleteCells.entireColumn' },
];

/**
 * Excel's suggestion: entire rows / columns for whole-row / whole-column
 * selections, a sideways shift for selections taller than they are wide,
 * else a vertical shift.
 */
function defaultChoice<S extends string>(
	store: ViewerStore,
	sideways: S,
	vertical: S,
): S | 'rows' | 'cols' {
	const sel = store.selectionOf();
	const axis = wholeAxisOfRanges(sel.ranges);
	if (axis) {
		return axis;
	}
	const r = currentRange(sel);
	return r.bottom - r.top > r.right - r.left ? sideways : vertical;
}

function ShiftChoiceDialog<T extends string>({
	title,
	choices,
	initial,
	onClose,
	onApply,
}: {
	title: string;
	choices: readonly Choice<T>[];
	initial: T;
	onClose: () => void;
	onApply: (choice: T) => void;
}) {
	const t = useT();
	const store = useViewerStore();
	const name = useId();
	const [choice, setChoice] = useState<T>(initial);
	const ok = () => {
		onClose();
		onApply(choice);
	};
	return (
		<FormDialog
			title={title}
			onClose={onClose}
			onOk={ok}
			okDisabled={store.readOnly}
			width={260}
			initialFocusSelector='input[type="radio"]:checked'
		>
			<GroupBox legend={title}>
				<div role='radiogroup' aria-label={title} className='flex flex-col gap-1.5 pt-1'>
					{choices.map((c) => (
						<Radio
							key={c.value}
							name={name}
							label={t(c.labelKey)}
							checked={choice === c.value}
							onChange={() => setChoice(c.value)}
						/>
					))}
				</div>
			</GroupBox>
		</FormDialog>
	);
}

/** The Insert dialog (kind `insertCells`). */
export function InsertCellsDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const [initial] = useState<InsertChoice>(() => defaultChoice(store, 'right', 'down'));
	return (
		<ShiftChoiceDialog
			title={t('dialog.insertCells.title')}
			choices={INSERT_CHOICES}
			initial={initial}
			onClose={onClose}
			onApply={(choice) => {
				if (choice === 'rows' || choice === 'cols') {
					insertCellsCmd(store, choice);
				} else {
					insertCellsShiftCmd(store, choice);
				}
			}}
		/>
	);
}

/** The Delete dialog (kind `deleteCells`). */
export function DeleteCellsDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const [initial] = useState<DeleteChoice>(() => defaultChoice(store, 'left', 'up'));
	return (
		<ShiftChoiceDialog
			title={t('dialog.deleteCells.title')}
			choices={DELETE_CHOICES}
			initial={initial}
			onClose={onClose}
			onApply={(choice) => {
				if (choice === 'rows' || choice === 'cols') {
					deleteCellsCmd(store, choice);
				} else {
					deleteCellsShiftCmd(store, choice);
				}
			}}
		/>
	);
}
