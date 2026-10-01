import type { MarginPreset } from 'xlsx-viewer-core';
import { getPrintArea, MARGIN_PRESETS, paperSize, paperSizeDetail } from 'xlsx-viewer-core';
import { File, Frame, PanelTop, RectangleVertical, Scan, SeparatorHorizontal } from 'lucide-react';

import { formatNumber, useLocale, useT, type TranslationKey } from '../../../../i18n';
import { useShell } from '../../../shell-context';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import {
	canRemoveBreakAtActive,
	marginsCmd,
	marginsMode,
	orientationCmd,
	pageBreakCmd,
	paperSizeCmd,
	printAreaCmd,
} from '../../print/page-setup-commands';
import { paperSizeOptions } from '../../print/PrintSettings';
import { MenuButton, RibbonButton } from '../../ui/buttons';
import type { MenuItem } from '../../ui/Menu';
import { RibbonGroup } from '../RibbonGroup';

const MARGIN_LABELS: Record<MarginPreset, TranslationKey> = {
	normal: 'pageLayout.marginsNormal',
	wide: 'pageLayout.marginsWide',
	narrow: 'pageLayout.marginsNarrow',
};

/** Page Layout > Page Setup: margins, orientation, size, print area, breaks and print titles. */
export function PageSetupGroup() {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const shell = useShell();
	// Re-render after edits and selection changes (the menus read the active sheet).
	const snap = useStoreSelector((s) => s);
	const sheetIndex = snap.ui.activeSheet;
	const sheet = store.workbook.sheets[sheetIndex];
	const ro = store.readOnly;
	const margins = marginsMode(sheet);
	const landscape = sheet.pageSetup?.orientation === 'landscape';
	const paper = paperSize(sheet.pageSetup?.paperSize).code;
	const hasArea = getPrintArea(store.workbook, sheetIndex).length > 0;
	const hasBreaks = (sheet.rowBreaks?.length ?? 0) + (sheet.colBreaks?.length ?? 0) > 0;

	const inches = (n: number) => formatNumber(n, locale);
	const marginItems: MenuItem[] = [
		...(['normal', 'wide', 'narrow'] as const).map((p) => {
			const m = MARGIN_PRESETS[p];
			return {
				label: t('pageLayout.marginPreset', {
					name: t(MARGIN_LABELS[p]),
					top: inches(m.top),
					bottom: inches(m.bottom),
					left: inches(m.left),
					right: inches(m.right),
				}),
				checked: margins === p,
				onSelect: () => marginsCmd(store, p),
			};
		}),
		{ separator: true },
		{
			label: t('print.customMargins'),
			onSelect: () => shell.openDialog('pageSetup', { tab: 'margins' }),
		},
	];
	const sizeItems: MenuItem[] = [
		...paperSizeOptions(paper).map((code) => {
			const size = paperSize(code);
			return {
				label: `${size.name}  (${paperSizeDetail(size)})`,
				checked: code === paper,
				onSelect: () => paperSizeCmd(store, code),
			};
		}),
		{ separator: true },
		{
			label: t('pageLayout.moreSizes'),
			onSelect: () => shell.openDialog('pageSetup', { tab: 'page' }),
		},
	];
	const areaItems: MenuItem[] = [
		{ label: t('pageLayout.setPrintArea'), onSelect: () => printAreaCmd(store, 'set') },
		{
			label: t('pageLayout.clearPrintArea'),
			disabled: !hasArea,
			onSelect: () => printAreaCmd(store, 'clear'),
		},
		...(hasArea
			? [{ label: t('pageLayout.addToPrintArea'), onSelect: () => printAreaCmd(store, 'add') }]
			: []),
	];
	const active = store.selectionOf(sheetIndex).active;
	const breakItems: MenuItem[] = [
		{
			label: t('pageLayout.insertBreak'),
			disabled: active.row === 0 && active.col === 0,
			onSelect: () => pageBreakCmd(store, 'insert'),
		},
		{
			label: t('pageLayout.removeBreak'),
			disabled: !canRemoveBreakAtActive(store),
			onSelect: () => pageBreakCmd(store, 'remove'),
		},
		{
			label: t('pageLayout.resetBreaks'),
			disabled: !hasBreaks,
			onSelect: () => pageBreakCmd(store, 'reset'),
		},
	];
	return (
		<RibbonGroup label={t('group.pageSetup')}>
			<MenuButton
				icon={<Frame size={15} />}
				label={t('cmd.margins')}
				disabled={ro}
				items={marginItems}
			/>
			<MenuButton
				icon={<RectangleVertical size={15} className={landscape ? 'rotate-90' : undefined} />}
				label={t('cmd.orientation')}
				disabled={ro}
				items={[
					{
						label: t('print.portrait'),
						checked: !landscape,
						onSelect: () => orientationCmd(store, 'portrait'),
					},
					{
						label: t('print.landscape'),
						checked: landscape,
						onSelect: () => orientationCmd(store, 'landscape'),
					},
				]}
			/>
			<MenuButton icon={<File size={15} />} label={t('cmd.size')} disabled={ro} items={sizeItems} />
			<MenuButton
				icon={<Scan size={15} />}
				label={t('cmd.printArea')}
				disabled={ro}
				items={areaItems}
			/>
			<MenuButton
				icon={<SeparatorHorizontal size={15} />}
				label={t('cmd.breaks')}
				disabled={ro}
				items={breakItems}
			/>
			<RibbonButton
				icon={<PanelTop size={15} />}
				label={t('cmd.printTitles')}
				showLabel
				disabled={ro}
				onClick={() => shell.openDialog('pageSetup', { tab: 'sheet' })}
			/>
		</RibbonGroup>
	);
}
