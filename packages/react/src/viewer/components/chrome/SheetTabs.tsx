import { resolveColorCss } from 'xlsx-viewer-core';
import { ChevronLeft, ChevronRight, Plus } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';

import { useT } from '../../../i18n';
import {
	deleteSheetCmd,
	duplicateSheetCmd,
	hideSheetCmd,
	insertSheetCmd,
	moveSheetCmd,
	renameSheetCmd,
	setTabColorCmd,
} from '../../commands/sheet-commands';
import { useReturnToGrid } from '../../shell-context';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { ColorPicker } from '../ui/ColorPicker';
import { MenuList, type MenuItem } from '../ui/Menu';
import { Popover } from '../ui/Popover';

/** Sheet tabs with add, rename (double-click), drag to reorder and a context menu. */
export function SheetTabs() {
	const t = useT();
	const store = useViewerStore();
	const returnToGrid = useReturnToGrid();
	const active = useStoreSelector((s) => s.ui.activeSheet);
	const sheets = useStoreSelector(
		(s, st) => (
			void s.docVersion,
			st.workbook.sheets.map((sh) => ({
				id: sh.id,
				name: sh.name,
				state: sh.state,
				tab: sh.tabColor,
			}))
		),
		(a, b) =>
			a.length === b.length &&
			a.every(
				(x, i) =>
					x.id === b[i].id && x.name === b[i].name && x.state === b[i].state && x.tab === b[i].tab,
			),
	);
	const [renaming, setRenaming] = useState<number | null>(null);
	/** Escape: the rename input loses focus without renaming. */
	const cancelRename = useRef(false);
	const [menu, setMenu] = useState<{ index: number; x: number; y: number } | null>(null);
	const [dragFrom, setDragFrom] = useState<number | null>(null);
	const scrollRef = useRef<HTMLDivElement>(null);
	const theme = store.workbook.theme;

	useEffect(() => {
		scrollRef.current
			?.querySelector<HTMLElement>('[aria-selected="true"]')
			?.scrollIntoView({ block: 'nearest', inline: 'nearest' });
	}, [active]);

	const menuItems = (index: number): MenuItem[] => [
		{ label: t('sheet.insert'), onSelect: () => insertSheetCmd(store, index) },
		{
			label: t('sheet.delete'),
			danger: true,
			disabled: store.readOnly,
			onSelect: () => deleteSheetCmd(store, index),
		},
		{ label: t('sheet.rename'), disabled: store.readOnly, onSelect: () => setRenaming(index) },
		{
			label: t('sheet.duplicate'),
			disabled: store.readOnly,
			onSelect: () => duplicateSheetCmd(store, index),
		},
		{ separator: true },
		{
			label: t('sheet.moveLeft'),
			disabled: store.readOnly || index === 0,
			onSelect: () => moveSheetCmd(store, index, index - 1),
		},
		{
			label: t('sheet.moveRight'),
			disabled: store.readOnly || index === sheets.length - 1,
			onSelect: () => moveSheetCmd(store, index, index + 1),
		},
		{
			label: t('sheet.tabColor'),
			disabled: store.readOnly,
			items: [
				{
					render: (close) => (
						<ColorPicker
							theme={theme}
							noneLabel={t('color.automatic')}
							close={close}
							onPick={(c) => setTabColorCmd(store, index, c ?? undefined)}
						/>
					),
				},
			],
		},
		{ separator: true },
		{
			label: t('sheet.hide'),
			disabled: store.readOnly,
			onSelect: () => hideSheetCmd(store, index, true),
		},
		{
			label: t('sheet.unhide'),
			disabled: store.readOnly || !sheets.some((s) => s.state === 'hidden'),
			items: sheets
				.map((s, i) => ({ s, i }))
				.filter(({ s }) => s.state === 'hidden')
				.map(({ s, i }) => ({ label: s.name, onSelect: () => hideSheetCmd(store, i, false) })),
		},
	];

	return (
		<div className='xv-sheet-tabs flex h-[30px] shrink-0 items-stretch border-t border-border bg-card text-[12.5px]'>
			<div className='flex items-center px-1 text-muted-foreground'>
				<button
					type='button'
					aria-label={t('sheet.prev')}
					className='rounded-sm p-1 hover:bg-accent'
					onClick={() => scrollRef.current?.scrollBy({ left: -120, behavior: 'smooth' })}
				>
					<ChevronLeft size={14} />
				</button>
				<button
					type='button'
					aria-label={t('sheet.next')}
					className='rounded-sm p-1 hover:bg-accent'
					onClick={() => scrollRef.current?.scrollBy({ left: 120, behavior: 'smooth' })}
				>
					<ChevronRight size={14} />
				</button>
			</div>
			<div
				ref={scrollRef}
				role='tablist'
				aria-label={t('sheet.tabList')}
				className='flex min-w-0 items-stretch overflow-x-auto [scrollbar-width:none]'
			>
				{sheets.map((s, i) => {
					if (s.state !== 'visible') {
						return null;
					}
					const color = s.tab ? resolveColorCss(s.tab, { theme }) : undefined;
					const selected = i === active;
					return (
						<div
							key={s.id}
							role='tab'
							aria-selected={selected}
							tabIndex={selected ? 0 : -1}
							draggable={!store.readOnly && renaming === null}
							className={`relative flex min-w-[64px] cursor-default select-none items-center justify-center border-r border-border px-3 ${
								selected
									? 'bg-background font-semibold text-primary'
									: 'text-foreground hover:bg-accent'
							} ${dragFrom !== null && dragFrom !== i ? 'data-[over=true]:shadow-[inset_2px_0_0_var(--xlsx-primary)]' : ''}`}
							onClick={(e) => {
								store.setUi({ activeSheet: i, edit: null });
								if (renaming !== i) {
									returnToGrid(e.currentTarget);
								}
							}}
							onDoubleClick={() => !store.readOnly && setRenaming(i)}
							onContextMenu={(e) => {
								e.preventDefault();
								store.setUi({ activeSheet: i });
								setMenu({ index: i, x: e.clientX, y: e.clientY });
							}}
							onDragStart={(e) => {
								setDragFrom(i);
								e.dataTransfer.effectAllowed = 'move';
							}}
							onDragOver={(e) => {
								if (dragFrom !== null) {
									e.preventDefault();
								}
							}}
							onDrop={(e) => {
								e.preventDefault();
								if (dragFrom !== null) {
									moveSheetCmd(store, dragFrom, i);
								}
								setDragFrom(null);
							}}
							onDragEnd={() => setDragFrom(null)}
						>
							{renaming === i ? (
								<input
									// biome-ignore lint: renaming starts on explicit user action
									autoFocus
									defaultValue={s.name}
									aria-label={t('sheet.rename')}
									className='w-[120px] rounded-sm border border-primary bg-card px-1 text-[12.5px] text-foreground outline-none'
									onFocus={(e) => e.currentTarget.select()}
									onBlur={(e) => {
										const name = e.currentTarget.value;
										if (!cancelRename.current && name.trim() !== s.name) {
											const err = renameSheetCmd(store, i, name);
											if (err) {
												store.setUi({ message: err });
											}
										}
										cancelRename.current = false;
										setRenaming(null);
									}}
									onKeyDown={(e) => {
										if (e.key === 'Enter') {
											returnToGrid(e.currentTarget); // blurring commits the name
										} else if (e.key === 'Escape') {
											cancelRename.current = true;
											returnToGrid(e.currentTarget);
										}
									}}
								/>
							) : (
								<span className='whitespace-nowrap'>{s.name}</span>
							)}
							{color && (
								<span
									className='absolute bottom-0 left-1 right-1 h-[3px] rounded-t-sm'
									style={{ backgroundColor: color }}
								/>
							)}
							{selected && !color && (
								<span className='absolute bottom-0 left-2 right-2 h-[2px] rounded-t-sm bg-primary' />
							)}
						</div>
					);
				})}
			</div>
			{!store.readOnly && (
				<button
					type='button'
					aria-label={t('sheet.add')}
					title={t('ui.withShortcut', { label: t('sheet.add'), shortcut: 'Shift+F11' })}
					className='mx-1 my-auto rounded-full p-1 text-muted-foreground hover:bg-accent hover:text-foreground'
					onClick={() => insertSheetCmd(store)}
				>
					<Plus size={15} />
				</button>
			)}
			<div className='flex-1' />
			<Popover
				open={menu !== null}
				onClose={() => setMenu(null)}
				placement='point'
				point={menu ? { x: menu.x, y: menu.y } : undefined}
				role='menu'
			>
				{menu && <MenuList items={menuItems(menu.index)} onClose={() => setMenu(null)} />}
			</Popover>
		</div>
	);
}
