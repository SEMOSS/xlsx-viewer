import { useId, useState } from 'react';

import {
	applyDraft,
	draftFromSheet,
	validateDraft,
	type PageSetupDraft,
} from '../../../engine/print/page-setup-draft';
import { useT, type TranslationKey } from '../../../i18n';
import { useViewerStore } from '../../state/use-store';
import { DialogMessage, FormDialog, TabPanel, TabStrip, type TabDef } from './dialog-ui';
import { HeaderFooterTab } from './page-setup/HeaderFooterTab';
import { MarginsTab } from './page-setup/MarginsTab';
import { PageTab } from './page-setup/PageTab';
import { SheetTab } from './page-setup/SheetTab';
import type { DialogProps } from './types';

export type PageSetupTab = 'page' | 'margins' | 'headerFooter' | 'sheet';

const TABS: ReadonlyArray<{ id: PageSetupTab; labelKey: TranslationKey }> = [
	{ id: 'page', labelKey: 'dialog.pageSetup.tabPage' },
	{ id: 'margins', labelKey: 'dialog.pageSetup.tabMargins' },
	{ id: 'headerFooter', labelKey: 'dialog.pageSetup.tabHeaderFooter' },
	{ id: 'sheet', labelKey: 'dialog.pageSetup.tabSheet' },
];

function initialTab(payload: unknown): PageSetupTab {
	const tab = (payload as { tab?: unknown } | undefined)?.tab;
	return TABS.some((x) => x.id === tab) ? (tab as PageSetupTab) : 'page';
}

/**
 * Page Layout > Page Setup: page, margins, header / footer and sheet
 * settings of the active sheet, applied as one undoable step on OK.
 * `payload.tab` opens a tab (Custom Margins, Print Titles...).
 */
export function PageSetupDialog({ onClose, payload }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const idBase = useId();
	const sheetIndex = store.state.activeSheet;
	const sheet = store.workbook.sheets[sheetIndex];
	const [tab, setTab] = useState<PageSetupTab>(() => initialTab(payload));
	const [draft, setDraft] = useState<PageSetupDraft>(() =>
		draftFromSheet(store.workbook, sheetIndex, t),
	);
	const [error, setError] = useState<string | null>(null);
	const update = (patch: Partial<PageSetupDraft>) => {
		setDraft((d) => ({ ...d, ...patch }));
		setError(null);
	};
	const ok = () => {
		const problem = validateDraft(draft, sheet.name, t);
		if (problem) {
			setError(problem);
			return;
		}
		if (!store.readOnly) {
			try {
				store.transact(
					store.t('print.pageSetup'),
					(tx) => applyDraft(tx, store.workbook, sheetIndex, draft),
					{ autoRowHeight: false },
				);
			} catch (e) {
				setError((e as Error).message);
				return;
			}
		}
		onClose();
	};
	const tabs: TabDef<PageSetupTab>[] = TABS.map((x) => ({ id: x.id, label: t(x.labelKey) }));
	return (
		<FormDialog
			title={t('print.pageSetup')}
			onClose={onClose}
			onOk={ok}
			width={520}
			okDisabled={store.readOnly}
		>
			<TabStrip
				tabs={tabs}
				value={tab}
				onChange={setTab}
				label={t('print.pageSetup')}
				idBase={idBase}
			/>
			<TabPanel idBase={idBase} id={tab} className='min-h-[300px]'>
				{tab === 'page' && <PageTab draft={draft} update={update} />}
				{tab === 'margins' && <MarginsTab draft={draft} update={update} />}
				{tab === 'headerFooter' && (
					<HeaderFooterTab
						draft={draft}
						update={update}
						sheetName={sheet.name}
						fileName='Book1.xlsx'
					/>
				)}
				{tab === 'sheet' && <SheetTab draft={draft} update={update} />}
			</TabPanel>
			{error && <DialogMessage>{error}</DialogMessage>}
		</FormDialog>
	);
}
