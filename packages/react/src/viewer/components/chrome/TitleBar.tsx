import { Moon, Redo2, Save, Search, Sun, Undo2 } from 'lucide-react';

import { useT } from '../../../i18n';
import { useStoreSelector, useViewerStore } from '../../state/use-store';
import { RibbonButton } from '../ui/buttons';

interface TitleBarProps {
	fileName: string;
	mode: 'light' | 'dark';
	onToggleTheme?: () => void;
	onSave?: () => void;
	onSearch?: () => void;
	canEdit: boolean;
}

/** App logo: a green sheet with an "X", drawn inline so no asset is needed. */
function Logo() {
	return (
		<svg width='20' height='20' viewBox='0 0 24 24' aria-hidden>
			<rect x='3' y='2' width='18' height='20' rx='2.5' fill='#107c41' />
			<path d='M9 2v20M3 8h18M3 14h18' stroke='#33c481' strokeWidth='1' opacity='0.55' />
			<rect x='1.5' y='7' width='11' height='11' rx='1.5' fill='#185c37' />
			<path d='M4.5 9.5l5 6M9.5 9.5l-5 6' stroke='#fff' strokeWidth='1.7' strokeLinecap='round' />
		</svg>
	);
}

/** Top bar: logo, quick-access actions, file name and status, search and theme. */
export function TitleBar({
	fileName,
	mode,
	onToggleTheme,
	onSave,
	onSearch,
	canEdit,
}: TitleBarProps) {
	const t = useT();
	const store = useViewerStore();
	const history = useStoreSelector((s) => s.history);
	const dirty = useStoreSelector((s) => s.dirty);
	return (
		<header className='xv-titlebar flex h-9 shrink-0 items-center gap-1 border-b border-border bg-card px-2 text-foreground'>
			<div className='flex items-center gap-2 pr-1'>
				<Logo />
			</div>
			{canEdit && (
				<div className='flex items-center gap-0.5 border-l border-border pl-1'>
					{onSave && (
						<RibbonButton icon={<Save size={15} />} label={t('title.save')} onClick={onSave} />
					)}
					<RibbonButton
						icon={<Undo2 size={15} />}
						label={
							history.undo
								? `${t('title.undo')}${t('ui.labelSeparator')}${history.undo}`
								: t('title.undo')
						}
						disabled={!history.undo}
						onClick={() => store.undo()}
					/>
					<RibbonButton
						icon={<Redo2 size={15} />}
						label={
							history.redo
								? `${t('title.redo')}${t('ui.labelSeparator')}${history.redo}`
								: t('title.redo')
						}
						disabled={!history.redo}
						onClick={() => store.redo()}
					/>
				</div>
			)}
			<div className='flex min-w-0 flex-1 items-center justify-center gap-2 px-2'>
				<span className='truncate text-[13px] font-medium' title={fileName}>
					{fileName}
				</span>
				{canEdit ? (
					<span className='shrink-0 text-[12px] text-muted-foreground'>
						{dirty ? `· ${t('app.edited')}` : `· ${t('app.saved')}`}
					</span>
				) : (
					<span className='shrink-0 rounded-sm bg-muted px-1.5 text-[11px] text-muted-foreground'>
						{t('app.readOnly')}
					</span>
				)}
			</div>
			<button
				type='button'
				onClick={onSearch}
				className='hidden h-7 w-56 items-center gap-2 rounded-md border border-input bg-background px-2 text-[12px] text-muted-foreground hover:border-primary sm:flex'
				aria-label={t('title.search')}
			>
				<Search size={14} />
				<span>{t('title.search')}</span>
			</button>
			{onToggleTheme && (
				<RibbonButton
					icon={mode === 'dark' ? <Sun size={15} /> : <Moon size={15} />}
					label={mode === 'dark' ? t('cmd.lightMode') : t('cmd.darkMode')}
					onClick={onToggleTheme}
				/>
			)}
		</header>
	);
}
