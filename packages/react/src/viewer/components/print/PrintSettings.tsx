import type { Worksheet } from 'xlsx-viewer-core';
import { COMMON_PAPER_SIZES, PAPER_SIZES, paperSize, paperSizeDetail } from 'xlsx-viewer-core';
import { Printer } from 'lucide-react';
import { useId, type ReactNode } from 'react';

import { useT } from '../../../i18n';
import { useShell } from '../../shell-context';
import type { ViewerStore } from '../../state/viewer-store';
import {
	marginsCmd,
	marginsMode,
	orientationCmd,
	paperSizeCmd,
	scalingCmd,
	scalingMode,
	type ScalingMode,
} from './page-setup-commands';

export type PrintWhat = 'active' | 'workbook' | 'selection';

const SELECT =
	'w-full rounded-md border border-input bg-card px-2 py-1.5 text-[13px] text-foreground outline-none focus:border-primary disabled:opacity-50';

function Setting({ label, children }: { label: string; children: (id: string) => ReactNode }) {
	const id = useId();
	return (
		<div className='flex flex-col gap-1'>
			<label htmlFor={id} className='text-[12px] text-muted-foreground'>
				{label}
			</label>
			{children(id)}
		</div>
	);
}

/** Paper sizes for the Size lists: the common ones first, then the current one if missing. */
export function paperSizeOptions(current: number | undefined): number[] {
	const codes = [...COMMON_PAPER_SIZES];
	const code = paperSize(current).code;
	if (!codes.includes(code)) {
		codes.push(code);
	}
	return codes;
}

export interface PrintSettingsProps {
	store: ViewerStore;
	sheet: Worksheet;
	what: PrintWhat;
	onWhat: (what: PrintWhat) => void;
	ignorePrintArea: boolean;
	onIgnorePrintArea: (on: boolean) => void;
	copies: number;
	onCopies: (copies: number) => void;
	pageCount: number;
	onPrint: () => void;
}

/** The left column of File > Print: Print button, copies and the page settings of the active sheet. */
export function PrintSettings(p: PrintSettingsProps) {
	const t = useT();
	const shell = useShell();
	const { store, sheet } = p;
	const ro = store.readOnly;
	const orientation = sheet.pageSetup?.orientation === 'landscape' ? 'landscape' : 'portrait';
	const margins = marginsMode(sheet);
	const scaling = scalingMode(sheet);
	const paper = paperSize(sheet.pageSetup?.paperSize).code;
	return (
		<div className='flex w-[300px] shrink-0 flex-col gap-3 overflow-y-auto pr-2'>
			<h2 className='text-[22px] font-semibold'>{t('print.title')}</h2>
			<div className='flex items-end gap-3'>
				<button
					type='button'
					onClick={p.onPrint}
					disabled={p.pageCount === 0}
					className='flex h-[72px] w-[84px] flex-col items-center justify-center gap-1 rounded-md border border-border bg-card text-[13px] hover:bg-accent disabled:opacity-40'
				>
					<Printer size={26} className='text-primary' />
					{t('print.print')}
				</button>
				<Setting label={t('print.copies')}>
					{(id) => (
						<input
							id={id}
							type='number'
							min={1}
							max={999}
							value={p.copies}
							onChange={(e) => p.onCopies(Math.max(1, Math.min(999, Number(e.target.value) || 1)))}
							className={`${SELECT} w-[80px]`}
						/>
					)}
				</Setting>
			</div>
			{p.copies > 1 && <p className='text-[12px] text-muted-foreground'>{t('print.copiesNote')}</p>}
			<h3 className='mt-2 text-[15px] font-semibold'>{t('print.settings')}</h3>
			<Setting label={t('print.what')}>
				{(id) => (
					<select
						id={id}
						className={SELECT}
						value={p.what}
						onChange={(e) => p.onWhat(e.target.value as PrintWhat)}
					>
						<option value='active'>{t('print.activeSheets')}</option>
						<option value='workbook'>{t('print.entireWorkbook')}</option>
						<option value='selection'>{t('print.selection')}</option>
					</select>
				)}
			</Setting>
			<label className='flex items-center gap-2 text-[13px]'>
				<input
					type='checkbox'
					className='accent-[var(--xlsx-primary)]'
					checked={p.ignorePrintArea}
					disabled={p.what === 'selection'}
					onChange={(e) => p.onIgnorePrintArea(e.target.checked)}
				/>
				{t('print.ignorePrintArea')}
			</label>
			<Setting label={t('cmd.orientation')}>
				{(id) => (
					<select
						id={id}
						className={SELECT}
						value={orientation}
						disabled={ro}
						onChange={(e) => orientationCmd(store, e.target.value as 'portrait' | 'landscape')}
					>
						<option value='portrait'>{t('print.portrait')}</option>
						<option value='landscape'>{t('print.landscape')}</option>
					</select>
				)}
			</Setting>
			<Setting label={t('print.paperSize')}>
				{(id) => (
					<select
						id={id}
						className={SELECT}
						value={paper}
						disabled={ro}
						onChange={(e) => paperSizeCmd(store, Number(e.target.value))}
					>
						{paperSizeOptions(paper).map((code) => {
							const size = PAPER_SIZES.find((s) => s.code === code) ?? paperSize(code);
							return (
								<option key={code} value={code}>
									{`${size.name} (${paperSizeDetail(size)})`}
								</option>
							);
						})}
					</select>
				)}
			</Setting>
			<Setting label={t('cmd.margins')}>
				{(id) => (
					<select
						id={id}
						className={SELECT}
						value={margins}
						disabled={ro}
						onChange={(e) => {
							const v = e.target.value;
							if (v === 'dialog') {
								shell.openDialog('pageSetup', { tab: 'margins' });
							} else if (v === 'normal' || v === 'wide' || v === 'narrow') {
								marginsCmd(store, v);
							}
						}}
					>
						<option value='normal'>{t('print.marginsNormal')}</option>
						<option value='wide'>{t('print.marginsWide')}</option>
						<option value='narrow'>{t('print.marginsNarrow')}</option>
						{margins === 'custom' && <option value='custom'>{t('print.marginsCustom')}</option>}
						<option value='dialog'>{t('print.customMargins')}</option>
					</select>
				)}
			</Setting>
			<Setting label={t('print.scaling')}>
				{(id) => (
					<select
						id={id}
						className={SELECT}
						value={scaling}
						disabled={ro}
						onChange={(e) => {
							const v = e.target.value as ScalingMode | 'dialog';
							if (v === 'dialog') {
								shell.openDialog('pageSetup', { tab: 'page' });
							} else if (v !== 'custom') {
								scalingCmd(store, v);
							}
						}}
					>
						<option value='none'>{t('print.noScaling')}</option>
						<option value='fitSheet'>{t('print.fitSheet')}</option>
						<option value='fitColumns'>{t('print.fitColumns')}</option>
						<option value='fitRows'>{t('print.fitRows')}</option>
						{scaling === 'custom' && <option value='custom'>{t('print.customScaling')}</option>}
						<option value='dialog'>{t('print.customScalingOptions')}</option>
					</select>
				)}
			</Setting>
			<button
				type='button'
				className='self-start text-[13px] text-primary underline-offset-2 hover:underline disabled:opacity-40'
				disabled={ro}
				onClick={() => shell.openDialog('pageSetup')}
			>
				{t('print.pageSetup')}
			</button>
		</div>
	);
}
