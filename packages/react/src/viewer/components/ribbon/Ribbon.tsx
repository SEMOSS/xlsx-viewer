import { tableIndexAt } from 'xlsx-viewer-core';
import { useEffect } from 'react';

import { useT, type TranslationKey } from '../../../i18n';
import type { RibbonTabId } from '../../state/types';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { ChartDesignTab } from './chart/ChartDesignTab';
import { DataTab } from './data/DataTab';
import { FormulasTab } from './formulas/FormulasTab';
import { AlignmentGroup } from './home/AlignmentGroup';
import { CellsGroup, EditingGroup } from './home/CellsEditingGroups';
import { ClipboardGroup } from './home/ClipboardGroup';
import { FontGroup } from './home/FontGroup';
import { NumberGroup } from './home/NumberGroup';
import { StylesGroup } from './home/StylesGroup';
import { InsertTab } from './insert/InsertTab';
import { PageLayoutTab } from './page-layout/PageLayoutTab';
import { ReviewTab } from './review/ReviewTab';
import { TableDesignTab } from './table/TableDesignTab';
import { ViewTab } from './view/ViewTab';

const TABS: Array<{ id: RibbonTabId; key: TranslationKey }> = [
	{ id: 'home', key: 'tab.home' },
	{ id: 'insert', key: 'tab.insert' },
	{ id: 'pageLayout', key: 'tab.pageLayout' },
	{ id: 'formulas', key: 'tab.formulas' },
	{ id: 'data', key: 'tab.data' },
	{ id: 'review', key: 'tab.review' },
	{ id: 'view', key: 'tab.view' },
];

function HomeTab() {
	return (
		<>
			<ClipboardGroup />
			<FontGroup />
			<AlignmentGroup />
			<NumberGroup />
			<StylesGroup />
			<CellsGroup />
			<EditingGroup />
		</>
	);
}

/** The tab strip plus the active tab's command groups. */
export function Ribbon({ onOpenBackstage }: { onOpenBackstage: () => void }) {
	const t = useT();
	const store = useViewerStore();
	const selectedTab = useStoreSelector((s) => s.ui.ribbonTab);
	const inTable = useStoreSelector((snap, st) => {
		void snap.docVersion;
		const { row, col } = st.selectionOf(snap.ui.activeSheet).active;
		return tableIndexAt(st.workbook, snap.ui.activeSheet, row, col) >= 0;
	});
	const chartSelected = useStoreSelector((snap, st) => {
		void snap.docVersion;
		const id = snap.ui.selectedObject;
		return (
			id !== null &&
			Boolean(
				st.workbook.sheets[snap.ui.activeSheet]?.drawings.some(
					(d) => d.id === id && d.kind === 'chart',
				),
			)
		);
	});
	// Contextual tabs disappear when the active cell leaves the table or the chart is deselected.
	const contextLost =
		(selectedTab === 'tableDesign' && !inTable) ||
		(selectedTab === 'chartDesign' && !chartSelected);
	const tab = contextLost ? 'home' : selectedTab;
	useEffect(() => {
		if (contextLost) {
			store.setUi({ ribbonTab: 'home' });
		}
	}, [contextLost, store]);
	return (
		<div className='xv-ribbon shrink-0 border-b border-border bg-card'>
			<div
				role='tablist'
				aria-label={t('ribbon.label')}
				className='flex h-[34px] items-end gap-0.5 px-2'
			>
				<button
					type='button'
					className='mr-1 rounded-t-sm px-3 pb-2 pt-1.5 text-[13px] font-semibold text-primary hover:bg-accent'
					onClick={onOpenBackstage}
				>
					{t('tab.file')}
				</button>
				{TABS.map((tb) => (
					<button
						key={tb.id}
						type='button'
						role='tab'
						aria-selected={tab === tb.id}
						className={`relative px-3 pb-2 pt-1.5 text-[13px] whitespace-nowrap transition-colors hover:text-foreground ${
							tab === tb.id
								? 'font-semibold text-foreground after:absolute after:bottom-0 after:left-3 after:right-3 after:h-[3px] after:rounded-t-sm after:bg-primary'
								: 'text-muted-foreground'
						}`}
						onClick={() => store.setUi({ ribbonTab: tb.id })}
					>
						{t(tb.key)}
					</button>
				))}
				{inTable && (
					<button
						type='button'
						role='tab'
						aria-selected={tab === 'tableDesign'}
						className={`relative px-3 pb-2 pt-1.5 text-[13px] whitespace-nowrap text-primary transition-colors ${
							tab === 'tableDesign'
								? 'font-semibold after:absolute after:bottom-0 after:left-3 after:right-3 after:h-[3px] after:rounded-t-sm after:bg-primary'
								: 'hover:text-foreground'
						}`}
						onClick={() => store.setUi({ ribbonTab: 'tableDesign' })}
					>
						{t('tab.tableDesign')}
					</button>
				)}
				{chartSelected && (
					<button
						type='button'
						role='tab'
						aria-selected={tab === 'chartDesign'}
						className={`relative px-3 pb-2 pt-1.5 text-[13px] whitespace-nowrap text-primary transition-colors ${
							tab === 'chartDesign'
								? 'font-semibold after:absolute after:bottom-0 after:left-3 after:right-3 after:h-[3px] after:rounded-t-sm after:bg-primary'
								: 'hover:text-foreground'
						}`}
						onClick={() => store.setUi({ ribbonTab: 'chartDesign' })}
					>
						{t('chart.tab')}
					</button>
				)}
			</div>
			<div
				role='tabpanel'
				className='xv-ribbon-panel flex h-[54px] items-stretch overflow-x-auto overflow-y-hidden border-t border-border px-1 py-1'
			>
				{tab === 'home' && <HomeTab />}
				{tab === 'insert' && <InsertTab />}
				{tab === 'pageLayout' && <PageLayoutTab />}
				{tab === 'formulas' && <FormulasTab />}
				{tab === 'data' && <DataTab />}
				{tab === 'review' && <ReviewTab />}
				{tab === 'view' && <ViewTab />}
				{tab === 'tableDesign' && <TableDesignTab />}
				{tab === 'chartDesign' && <ChartDesignTab />}
			</div>
		</div>
	);
}
