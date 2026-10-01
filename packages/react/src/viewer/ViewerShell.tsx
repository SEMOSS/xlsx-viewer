import { getFormulaReferences } from 'xlsx-viewer-core';
import { useCallback, useEffect, useMemo, useRef, useState } from 'react';

import type { GridCommand } from '../engine/keymap';
import { currentRange } from '../engine/selection/selection';
import { paintFormats } from './commands/format-commands';
import { showRibbonTabFor } from './commands/shell-commands';
import {
	deleteCellsCmd,
	insertCellsCmd,
	toggleAutoFilter,
	wholeAxisOfSelection,
} from './commands/structure-commands';
import { insertSheetCmd } from './commands/sheet-commands';
import { hideColumns, hideRows } from './commands/layout-commands';
import { insertChart, insertPicture } from './commands/object-commands';
import { insertTableCmd } from './commands/table-commands';
import { zoomToSelection } from './commands/view-commands';
import { SheetGeometry } from '../engine/geometry/sheet-geometry';
import { Backstage } from './components/chrome/Backstage';
import { FormulaBar } from './components/chrome/FormulaBar';
import { GridContextMenu, type ContextMenuState } from './components/chrome/GridContextMenu';
import { SheetTabs } from './components/chrome/SheetTabs';
import { StatusBar } from './components/chrome/StatusBar';
import { TitleBar } from './components/chrome/TitleBar';
import { DialogHost, isDialogAvailable } from './components/dialogs/DialogHost';
import { Grid } from './components/grid/Grid';
import type { RefHighlight } from './components/grid/SelectionOverlay';
import { Ribbon } from './components/ribbon/Ribbon';
import { RibbonHostContext, type RibbonHost } from './components/ribbon/ribbon-host';
import { PortalContext } from './components/ui/Popover';
import { downloadWorkbook, printSheet } from './download';
import { focusGridIn, keepGridFocus } from './focus';
import { ShellContext, type DialogKind, type ShellActions } from './shell-context';
import { useStoreSelector, useViewerStore } from './state/use-store';

/** Colors Excel uses for formula references, in order of appearance. */
export const REFERENCE_COLORS = [
	'#3c78d8',
	'#e06666',
	'#8e7cc3',
	'#6aa84f',
	'#c27ba0',
	'#e69138',
	'#45818e',
	'#a64d79',
];

interface ViewerShellProps {
	fileName: string;
	canEdit: boolean;
	mode: 'light' | 'dark';
	onToggleTheme: () => void;
	onSave: () => Promise<void>;
	onOpenFile?: () => void;
	onNewWorkbook?: () => void;
	authorName: string;
	showRibbon: boolean;
	showFormulaBar: boolean;
	showSheetTabs: boolean;
	showStatusBar: boolean;
	showTitleBar: boolean;
	registerFocus: (fn: () => void) => void;
}

/** The viewer layout: title bar, ribbon, formula bar, grid, sheet tabs, status bar and overlays. */
export function ViewerShell(props: ViewerShellProps) {
	const store = useViewerStore();
	const [portalHost, setPortalHost] = useState<HTMLElement | null>(null);
	const [backstage, setBackstage] = useState(false);
	const [menu, setMenu] = useState<ContextMenuState | null>(null);
	const [formulaBar, setFormulaBar] = useState(props.showFormulaBar);
	const rootRef = useRef<HTMLDivElement>(null);
	const edit = useStoreSelector((s) => s.ui.edit);
	const activeSheet = useStoreSelector((s) => s.ui.activeSheet);
	const painter = useStoreSelector((s) => s.ui.formatPainter);
	const selection = useStoreSelector((s, st) => st.selectionOf(s.ui.activeSheet));

	useEffect(() => setFormulaBar(props.showFormulaBar), [props.showFormulaBar]);

	const focusGrid = useCallback(() => focusGridIn(rootRef.current), []);

	useEffect(() => {
		props.registerFocus(focusGrid);
	}, [props, focusGrid]);

	// Format painter: apply when the user finishes a selection.
	const lastSel = useRef(selection);
	useEffect(() => {
		if (painter && lastSel.current !== selection) {
			const target = currentRange(selection);
			const up = () => {
				paintFormats(store, painter, currentRange(store.selectionOf()));
				if (painter.mode === 'once') {
					store.setUi({ formatPainter: null });
				}
				window.removeEventListener('pointerup', up);
			};
			window.addEventListener('pointerup', up, { once: true });
			void target;
		}
		lastSel.current = selection;
	}, [selection, painter, store]);

	const openDialog = useCallback(
		(kind: DialogKind, payload?: unknown) => {
			if (!isDialogAvailable(kind)) {
				store.setUi({ message: store.t('message.dialogUnavailable') });
				return;
			}
			store.setUi({ dialog: { kind, payload } });
		},
		[store],
	);

	const shell: ShellActions = useMemo(
		() => ({
			openDialog,
			save: () => void props.onSave(),
			download: (format) =>
				void downloadWorkbook(store.workbook, store.state.activeSheet, props.fileName, format),
			openFile: props.onOpenFile,
			newWorkbook: props.onNewWorkbook,
			print: () => printSheet(store.workbook, store.state.activeSheet),
			notify: (message) => store.setUi({ message }),
			toggleTheme: props.onToggleTheme,
			themeMode: props.mode,
			focusGrid,
		}),
		[openDialog, props, store, focusGrid],
	);

	const pictureInput = useRef<HTMLInputElement>(null);
	const ribbonHost: RibbonHost = useMemo(
		() => ({
			toggleFormulaBar: () => setFormulaBar((v) => !v),
			formulaBarVisible: formulaBar,
			insertPicture: () => pictureInput.current?.click(),
			insertChart: (type) => {
				try {
					insertChart(store, type);
				} catch (e) {
					store.setUi({ message: (e as Error).message });
				}
			},
			insertTable: () => insertTableCmd(store),
			zoomToSelection: () => {
				const grid = rootRef.current?.querySelector('.xv-grid');
				if (!grid) {
					return;
				}
				const sheet = store.activeSheet;
				const zoom = (sheet.view.zoomScale || 100) / 100;
				const r = currentRange(store.selectionOf());
				const g = new SheetGeometry(sheet, { zoom });
				const rect = g.rangeRect(r);
				zoomToSelection(
					store,
					{ width: grid.clientWidth - 50, height: grid.clientHeight - 30 },
					rect,
				);
			},
		}),
		[formulaBar, store],
	);

	const onGridAction = useCallback(
		(action: GridCommand['id']) => {
			switch (action) {
				case 'save':
					void props.onSave();
					break;
				case 'open':
					props.onOpenFile?.();
					break;
				case 'print':
					printSheet(store.workbook, store.state.activeSheet);
					break;
				case 'find':
					openDialog('find');
					break;
				case 'replace':
					openDialog('replace');
					break;
				case 'goTo':
					openDialog('goTo');
					break;
				case 'formatCells':
					openDialog('formatCells');
					break;
				case 'hyperlink':
					openDialog('hyperlink');
					break;
				case 'note':
					openDialog('note');
					break;
				case 'filter':
					toggleAutoFilter(store);
					break;
				case 'newSheet':
					insertSheetCmd(store);
					break;
				case 'unhideRows':
					hideRows(store, false);
					break;
				case 'unhideCols':
					hideColumns(store, false);
					break;
				case 'insertCells':
				case 'deleteCells': {
					const axis = wholeAxisOfSelection(store);
					if (axis) {
						(action === 'insertCells' ? insertCellsCmd : deleteCellsCmd)(store, axis);
					} else {
						openDialog(action);
					}
					break;
				}
				case 'table':
					insertTableCmd(store);
					break;
				default:
					showRibbonTabFor(store, action);
			}
		},
		[openDialog, props, store],
	);

	// Formula reference highlights while editing a formula.
	const refHighlights: RefHighlight[] | undefined = useMemo(() => {
		if (!edit || !edit.text.startsWith('=') || edit.sheet !== activeSheet) {
			return undefined;
		}
		try {
			const sheetName = store.workbook.sheets[edit.sheet].name;
			return getFormulaReferences(edit.text.slice(1), sheetName)
				.filter((r) => !r.sheet || r.sheet.toLowerCase() === sheetName.toLowerCase())
				.map((r) => ({
					range: r.range,
					color: REFERENCE_COLORS[r.colorIndex % REFERENCE_COLORS.length],
				}));
		} catch {
			return undefined;
		}
	}, [edit, activeSheet, store]);

	return (
		<PortalContext.Provider value={portalHost}>
			<ShellContext.Provider value={shell}>
				<RibbonHostContext.Provider value={ribbonHost}>
					<div
						ref={rootRef}
						className='xv-shell relative flex h-full w-full flex-col overflow-hidden bg-background text-foreground'
						onMouseDown={(e) => keepGridFocus(e.nativeEvent, e.currentTarget, !!store.state.edit)}
					>
						{props.showTitleBar && (
							<TitleBar
								fileName={props.fileName}
								mode={props.mode}
								onToggleTheme={props.onToggleTheme}
								onSave={props.canEdit ? () => void props.onSave() : undefined}
								onSearch={() => openDialog('find')}
								canEdit={props.canEdit}
							/>
						)}
						{props.showRibbon && <Ribbon onOpenBackstage={() => setBackstage(true)} />}
						{formulaBar && <FormulaBar />}
						<Grid
							onAction={onGridAction}
							refHighlights={refHighlights}
							onContextMenu={(e) =>
								setMenu({
									x: e.clientX,
									y: e.clientY,
									target:
										e.hit?.kind === 'rowHeader'
											? 'rowHeader'
											: e.hit?.kind === 'colHeader'
												? 'colHeader'
												: 'cell',
								})
							}
						/>
						{props.showSheetTabs && <SheetTabs />}
						{props.showStatusBar && <StatusBar />}
						{backstage && (
							<Backstage fileName={props.fileName} onClose={() => setBackstage(false)} />
						)}
						<GridContextMenu state={menu} onClose={() => setMenu(null)} />
						<DialogHost author={props.authorName} />
						<div ref={setPortalHost} className='xv-portal' />
						<input
							ref={pictureInput}
							type='file'
							accept='image/png,image/jpeg,image/gif,image/bmp,image/webp,image/svg+xml'
							className='hidden'
							aria-hidden
							tabIndex={-1}
							onChange={(e) => {
								const file = e.target.files?.[0];
								e.target.value = '';
								if (file) {
									void insertPicture(store, file);
								}
							}}
						/>
					</div>
				</RibbonHostContext.Provider>
			</ShellContext.Provider>
		</PortalContext.Provider>
	);
}
