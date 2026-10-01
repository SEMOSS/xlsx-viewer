import { ListChecks, ListX } from 'lucide-react';
import { useEffect, useMemo, useState, type ReactNode } from 'react';
import type { RangeAddress } from 'xlsx-viewer-core';
import {
	cellKey,
	columnToLetters,
	formatValue,
	guessHeader,
	trimRangeToUsed,
} from 'xlsx-viewer-core';

import { useT } from '../../../i18n';
import { removeDuplicatesCmd } from '../../commands/data-tools-commands';
import { dataBlock } from '../../commands/structure-commands';
import { useViewerStore } from '../../state/use-store';
import type { ViewerStore } from '../../state/viewer-store';
import { Checkbox, FormDialog } from './dialog-ui';
import type { DialogProps } from './types';

function headerText(store: ViewerStore, row: number, col: number): string {
	const cell = store.activeSheet.cells.get(cellKey(row, col));
	if (!cell || cell.value === null) {
		return '';
	}
	const numFmt = store.styles.resolve(cell.style).numFmt;
	return formatValue(cell.value, numFmt, { date1904: store.workbook.date1904 }).text.trim();
}

/** Show which rows take part, like Excel: the data rows below the header. */
function selectRows(store: ViewerStore, block: RangeAddress, hasHeader: boolean): void {
	const range = hasHeader && block.bottom > block.top ? { ...block, top: block.top + 1 } : block;
	const active = { row: range.top, col: range.left };
	store.setSelection({ active, anchor: active, ranges: [range] });
}

const TOOL_BUTTON =
	'flex items-center gap-1.5 rounded-md border border-border bg-card px-2.5 py-1 text-[13px] ' +
	'hover:bg-accent';

const COLUMN_LIST =
	'flex max-h-[180px] flex-col gap-1 overflow-y-auto rounded-md border border-input bg-card ' +
	'px-2 py-1.5';

function ToolButton({
	icon,
	onClick,
	children,
}: {
	icon: ReactNode;
	onClick: () => void;
	children: string;
}) {
	return (
		<button type='button' className={TOOL_BUTTON} onClick={onClick}>
			{icon}
			{children}
		</button>
	);
}

/**
 * Data > Remove Duplicates: delete the rows of the selection (or of the
 * current region of a single cell) that repeat an earlier row in the
 * checked columns.
 */
export function RemoveDuplicatesDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const [block] = useState(() => trimRangeToUsed(store.activeSheet, dataBlock(store)));
	const [hasHeader, setHasHeader] = useState(() => guessHeader(store.activeSheet, block));
	const columns = useMemo(
		() => Array.from({ length: block.right - block.left + 1 }, (_, i) => block.left + i),
		[block],
	);
	const [checked, setChecked] = useState<ReadonlySet<number>>(() => new Set(columns));

	useEffect(() => {
		selectRows(store, block, hasHeader);
		// Only when the dialog opens; toggling the header option reselects explicitly.
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, []);

	const label = (col: number) =>
		(hasHeader && headerText(store, block.top, col)) ||
		t('dialog.removeDuplicates.column', { column: columnToLetters(col) });
	const toggleHeaders = (value: boolean) => {
		setHasHeader(value);
		selectRows(store, block, value);
	};
	const toggle = (col: number, on: boolean) => {
		const next = new Set(checked);
		if (on) {
			next.add(col);
		} else {
			next.delete(col);
		}
		setChecked(next);
	};
	const ok = () => {
		onClose();
		removeDuplicatesCmd(store, block, {
			columns: columns.filter((c) => checked.has(c)),
			hasHeader,
		});
	};

	return (
		<FormDialog
			title={t('cmd.removeDuplicates')}
			onClose={onClose}
			onOk={ok}
			okDisabled={checked.size === 0 || store.readOnly}
			width={440}
		>
			<p>{t('dialog.removeDuplicates.intro')}</p>
			<div className='flex flex-wrap items-center gap-2'>
				<ToolButton icon={<ListChecks size={15} />} onClick={() => setChecked(new Set(columns))}>
					{t('dialog.removeDuplicates.selectAll')}
				</ToolButton>
				<ToolButton icon={<ListX size={15} />} onClick={() => setChecked(new Set())}>
					{t('dialog.removeDuplicates.unselectAll')}
				</ToolButton>
				<Checkbox
					className='ml-auto'
					label={t('dialog.removeDuplicates.hasHeaders')}
					checked={hasHeader}
					onChange={toggleHeaders}
				/>
			</div>
			<div className='flex flex-col gap-1'>
				<span id='xv-remove-duplicates-columns'>{t('dialog.removeDuplicates.columns')}</span>
				<div role='group' aria-labelledby='xv-remove-duplicates-columns' className={COLUMN_LIST}>
					{columns.map((col) => (
						<Checkbox
							key={col}
							label={label(col)}
							checked={checked.has(col)}
							onChange={(on) => toggle(col, on)}
						/>
					))}
				</div>
			</div>
		</FormDialog>
	);
}
