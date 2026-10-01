import type { CellValue } from 'xlsx-viewer-core';
import { customNumberFormatList } from 'xlsx-viewer-core';
import { useMemo, useState } from 'react';

import { useT, type TranslationKey } from '../../../i18n';
import {
	applyFormatCells,
	formatCellsChanges,
	MIXED,
	readFormatCells,
	type FormatCellsValues,
} from '../../commands/format-cells-commands';
import { useViewerStore } from '../../state/use-store';
import { FormDialog, TabPanel, TabStrip, type TabDef } from './dialog-ui';
import { AlignmentTab } from './format-cells/AlignmentTab';
import { BorderTab, type BorderDraft } from './format-cells/BorderTab';
import {
	FillTab,
	fillDraftFrom,
	fillFromDraft,
	ProtectionTab,
	type FillDraft,
} from './format-cells/FillTab';
import { FontTab } from './format-cells/FontTab';
import { NumberTab } from './format-cells/NumberTab';
import { numberDraftCode, numberDraftFrom, type NumberDraft } from './format-cells/number-draft';
import type { DialogProps } from './types';

export type FormatCellsTab = 'number' | 'alignment' | 'font' | 'border' | 'fill' | 'protection';

const TABS: ReadonlyArray<{ id: FormatCellsTab; labelKey: TranslationKey }> = [
	{ id: 'number', labelKey: 'dialog.formatCells.tab.number' },
	{ id: 'alignment', labelKey: 'dialog.formatCells.tab.alignment' },
	{ id: 'font', labelKey: 'dialog.formatCells.tab.font' },
	{ id: 'border', labelKey: 'dialog.formatCells.tab.border' },
	{ id: 'fill', labelKey: 'dialog.formatCells.tab.fill' },
	{ id: 'protection', labelKey: 'dialog.formatCells.tab.protection' },
];

/** Excel reopens Format Cells on the tab used last. */
let lastTab: FormatCellsTab = 'number';

function payloadTab(payload: unknown): FormatCellsTab | undefined {
	const tab = (payload as { tab?: unknown } | undefined)?.tab;
	return TABS.some((def) => def.id === tab) ? (tab as FormatCellsTab) : undefined;
}

/** Format Cells (Ctrl+1): every change is applied to the selection in one undoable step. */
export function FormatCellsDialog({ onClose, payload }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const ctx = useMemo(() => readFormatCells(store), [store]);
	const theme = store.workbook.theme;
	const [tab, setTab] = useState<FormatCellsTab>(() => payloadTab(payload) ?? lastTab);
	const [draft, setDraft] = useState<FormatCellsValues>(ctx.values);
	const [number, setNumber] = useState<NumberDraft>(() => numberDraftFrom(ctx.activeFormat));
	const [numberTouched, setNumberTouched] = useState(false);
	const [border, setBorder] = useState<BorderDraft>({
		change: {},
		lineStyle: 'thin',
		lineColor: undefined,
	});
	const initialFill = ctx.values.fill === MIXED ? undefined : ctx.values.fill;
	const [fill, setFill] = useState<FillDraft>(() => fillDraftFrom(initialFill));
	const [fillTouched, setFillTouched] = useState(false);
	const customCodes = useMemo(() => customNumberFormatList(store.workbook.styles), [store]);
	const patch = (p: Partial<FormatCellsValues>) => setDraft((d) => ({ ...d, ...p }));

	const ok = () => {
		lastTab = tab;
		const values: FormatCellsValues = { ...draft };
		if (numberTouched) {
			values.numFmt = numberDraftCode(number);
		}
		if (fillTouched) {
			values.fill = fillFromDraft(fill);
		}
		applyFormatCells(store, formatCellsChanges(ctx.values, values, border.change));
		onClose();
	};

	const cancel = () => {
		lastTab = tab;
		onClose();
	};

	const idBase = 'xv-format-cells';
	const title = t('dialog.formatCells.title');
	const tabs: readonly TabDef<FormatCellsTab>[] = TABS.map((def) => ({
		id: def.id,
		label: t(def.labelKey),
	}));
	return (
		<FormDialog
			title={title}
			onClose={cancel}
			onOk={ok}
			okDisabled={store.readOnly}
			width={600}
			initialFocusSelector='[role="tab"][aria-selected="true"]'
		>
			<TabStrip tabs={tabs} value={tab} onChange={setTab} label={title} idBase={idBase} />
			<TabPanel idBase={idBase} id={tab} className='min-h-[372px]'>
				{tab === 'number' && (
					<NumberTab
						draft={number}
						onChange={(d) => {
							setNumber(d);
							setNumberTouched(true);
						}}
						sampleValue={ctx.sampleValue as CellValue}
						date1904={store.workbook.date1904}
						customCodes={customCodes}
					/>
				)}
				{tab === 'alignment' && <AlignmentTab draft={draft} onChange={patch} />}
				{tab === 'font' && (
					<FontTab
						draft={draft}
						onChange={patch}
						theme={theme}
						normalFont={store.workbook.styles.fonts[0] ?? {}}
					/>
				)}
				{tab === 'border' && (
					<BorderTab
						initial={ctx.borders}
						draft={border}
						onChange={setBorder}
						multiRow={ctx.multiRow}
						multiCol={ctx.multiCol}
						theme={theme}
					/>
				)}
				{tab === 'fill' && (
					<FillTab
						draft={fill}
						onChange={(d) => {
							setFill(d);
							setFillTouched(true);
						}}
						theme={theme}
						gradient={initialFill?.type === 'gradient' && !fillTouched}
					/>
				)}
				{tab === 'protection' && (
					<ProtectionTab
						locked={draft.locked === true}
						hidden={draft.hidden === true}
						lockedMixed={draft.locked === MIXED}
						hiddenMixed={draft.hidden === MIXED}
						onChange={(p) => patch(p)}
					/>
				)}
			</TabPanel>
		</FormDialog>
	);
}
