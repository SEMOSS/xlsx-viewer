import {
	columnWidthToPixels,
	getColInfo,
	MAX_COL_INDEX,
	MAX_ROW_INDEX,
	setColumnWidth,
	setRowHeight,
	storedWidthToDisplayChars,
	displayCharsToStoredWidth,
} from 'xlsx-viewer-core';
import { useState } from 'react';

import { formatPercent, useLocale, useT } from '../../../i18n';
import { currentRange } from '../../../engine/selection/selection';
import { renameSheetCmd } from '../../commands/sheet-commands';
import { goToReference } from '../../commands/selection-commands';
import { setZoom } from '../../commands/view-commands';
import { useViewerStore } from '../../state/use-store';
import { Dialog, DialogButton } from '../ui/Dialog';

const field =
	'w-full rounded-md border border-input bg-card px-2 py-1.5 text-[13px] text-foreground outline-none focus:border-primary';

export function GoToDialog({ onClose }: { onClose: () => void }) {
	const t = useT();
	const store = useViewerStore();
	const [ref, setRef] = useState('');
	const names = store.workbook.definedNames.filter((n) => !n.hidden && !n.builtin);
	const go = () => {
		if (goToReference(store, ref)) {
			onClose();
		} else {
			store.setUi({ message: t('message.invalidReferenceAt', { ref }) });
		}
	};
	return (
		<Dialog
			title={t('dialog.goTo.title')}
			onClose={onClose}
			footer={
				<>
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary onClick={go}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			{names.length > 0 && (
				<div className='mb-3 max-h-[160px] overflow-y-auto rounded-md border border-border'>
					{names.map((n) => (
						<button
							key={`${n.name}:${n.localSheet}`}
							type='button'
							className='block w-full px-2 py-1 text-left hover:bg-accent'
							onClick={() => setRef(n.name)}
							onDoubleClick={() => goToReference(store, n.name) && onClose()}
						>
							{n.name}
						</button>
					))}
				</div>
			)}
			<label className='flex flex-col gap-1'>
				<span>{t('dialog.goTo.reference')}</span>
				<input
					className={field}
					value={ref}
					onChange={(e) => setRef(e.target.value)}
					onKeyDown={(e) => e.key === 'Enter' && go()}
					placeholder={t('dialog.goTo.placeholder', { refs: 'B12, Sheet2!A1:C10' })}
				/>
			</label>
		</Dialog>
	);
}

export function SizeDialog({
	kind,
	onClose,
}: {
	kind: 'rowHeight' | 'columnWidth';
	onClose: () => void;
}) {
	const t = useT();
	const store = useViewerStore();
	const sheet = store.activeSheet;
	const r = currentRange(store.selectionOf());
	const initial =
		kind === 'rowHeight'
			? String(sheet.rows.get(r.top)?.height ?? sheet.defaultRowHeight)
			: String(
					storedWidthToDisplayChars(
						getColInfo(sheet, r.left)?.width ?? sheet.defaultColWidth ?? 9.140625,
					),
				);
	const [value, setValue] = useState(initial);
	const apply = () => {
		const n = Number(value);
		const max = kind === 'rowHeight' ? 409 : 255;
		if (!Number.isFinite(n) || n < 0 || n > max) {
			store.setUi({ message: t('message.numberBetween', { min: 0, max }) });
			return;
		}
		const sheetIndex = store.state.activeSheet;
		store.transact(store.t(kind === 'rowHeight' ? 'undo.rowHeight' : 'undo.columnWidth'), (tx) => {
			for (const range of store.selectionOf().ranges) {
				if (kind === 'rowHeight') {
					setRowHeight(
						tx,
						store.workbook,
						sheetIndex,
						range.top,
						Math.min(range.bottom, MAX_ROW_INDEX),
						n,
					);
				} else {
					setColumnWidth(
						tx,
						store.workbook,
						sheetIndex,
						range.left,
						Math.min(range.right, MAX_COL_INDEX),
						n === 0 ? 0 : displayCharsToStoredWidth(n),
					);
				}
			}
		});
		onClose();
	};
	void columnWidthToPixels;
	return (
		<Dialog
			title={t(kind === 'rowHeight' ? 'dialog.size.rowHeight' : 'dialog.size.columnWidth')}
			onClose={onClose}
			width={300}
			footer={
				<>
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary onClick={apply}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			<label className='flex items-center gap-2'>
				<span className='whitespace-nowrap'>
					{t(kind === 'rowHeight' ? 'dialog.size.rowHeightLabel' : 'dialog.size.columnWidthLabel')}
				</span>
				<input
					className={field}
					value={value}
					onChange={(e) => setValue(e.target.value)}
					onKeyDown={(e) => e.key === 'Enter' && apply()}
					inputMode='decimal'
				/>
			</label>
		</Dialog>
	);
}

export function ZoomDialog({ onClose }: { onClose: () => void }) {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const current = store.activeSheet.view.zoomScale || 100;
	const [value, setValue] = useState(String(current));
	const presets = [200, 100, 75, 50, 25];
	const apply = (z: number) => {
		setZoom(store, z);
		onClose();
	};
	return (
		<Dialog
			title={t('cmd.zoom')}
			onClose={onClose}
			width={260}
			footer={
				<>
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary onClick={() => apply(Number(value) || 100)}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			<div className='flex flex-col gap-1.5'>
				{presets.map((p) => (
					<label key={p} className='flex items-center gap-2'>
						<input
							type='radio'
							name='zoom'
							checked={Number(value) === p}
							onChange={() => setValue(String(p))}
						/>
						{formatPercent(p, locale)}
					</label>
				))}
				<label className='mt-1 flex items-center gap-2'>
					{t('dialog.zoom.custom')}
					<input
						className={`${field} w-20`}
						value={value}
						onChange={(e) => setValue(e.target.value)}
						inputMode='numeric'
					/>
					%
				</label>
			</div>
		</Dialog>
	);
}

export function RenameSheetDialog({ onClose }: { onClose: () => void }) {
	const t = useT();
	const store = useViewerStore();
	const index = store.state.activeSheet;
	const [name, setName] = useState(store.activeSheet.name);
	const [error, setError] = useState<string | null>(null);
	const apply = () => {
		const err = renameSheetCmd(store, index, name);
		if (err) {
			setError(err);
		} else {
			onClose();
		}
	};
	return (
		<Dialog
			title={t('cmd.renameSheet')}
			onClose={onClose}
			width={320}
			footer={
				<>
					<DialogButton onClick={onClose}>{t('dialog.cancel')}</DialogButton>
					<DialogButton primary onClick={apply}>
						{t('dialog.ok')}
					</DialogButton>
				</>
			}
		>
			<input
				className={field}
				value={name}
				maxLength={31}
				onChange={(e) => setName(e.target.value)}
				onKeyDown={(e) => e.key === 'Enter' && apply()}
				aria-label={t('dialog.renameSheet.sheetName')}
			/>
			{error && <p className='mt-2 text-[12px] text-destructive'>{error}</p>}
		</Dialog>
	);
}
