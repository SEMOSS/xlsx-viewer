import {
	ArrowLeft,
	Download,
	FileJson,
	FilePlus2,
	FileSpreadsheet,
	FileText,
	FolderOpen,
	Globe,
	Info,
	Printer,
	Save,
} from 'lucide-react';
import { useEffect, useRef, useState, type ReactNode } from 'react';

import { formatNumber, useLocale, useT } from '../../../i18n';
import { useViewerStore } from '../../state/use-store';
import { useShell } from '../../shell-context';
import { PrintView } from '../print/PrintView';
import { ProtectWorkbook } from './ProtectWorkbook';

function Row({
	icon,
	label,
	detail,
	onClick,
	disabled,
}: {
	icon: ReactNode;
	label: string;
	detail?: string;
	onClick?: () => void;
	disabled?: boolean;
}) {
	return (
		<button
			type='button'
			disabled={disabled}
			onClick={onClick}
			className='flex w-full items-center gap-3 rounded-md px-3 py-2.5 text-left hover:bg-accent disabled:opacity-40'
		>
			<span className='text-primary'>{icon}</span>
			<span className='flex flex-col'>
				<span className='text-[13.5px] font-medium text-foreground'>{label}</span>
				{detail && <span className='text-[12px] text-muted-foreground'>{detail}</span>}
			</span>
		</button>
	);
}

/** File menu (backstage view): info, save / download, export and print. */
export function Backstage({ fileName, onClose }: { fileName: string; onClose: () => void }) {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const shell = useShell();
	const wb = store.workbook;
	const props = wb.properties;
	const cellCount = wb.sheets.reduce((n, s) => n + s.cells.size, 0);
	const [printing, setPrinting] = useState(false);
	const run = (fn: () => void) => () => {
		onClose();
		fn();
	};
	// The backstage has the keyboard while it is open, then the grid gets it back.
	const backRef = useRef<HTMLButtonElement>(null);
	const shellRef = useRef(shell);
	shellRef.current = shell;
	useEffect(() => {
		backRef.current?.focus({ preventScroll: true });
		return () => shellRef.current.focusGrid?.();
	}, []);
	return (
		<div
			className='xv-backstage absolute inset-0 z-[500] flex bg-background text-foreground'
			onKeyDown={(e) => {
				if (e.key === 'Escape') {
					e.stopPropagation();
					onClose();
				}
			}}
		>
			<nav className='flex w-[220px] shrink-0 flex-col gap-1 bg-primary/95 p-3 text-primary-foreground'>
				<button
					ref={backRef}
					type='button'
					onClick={onClose}
					className='mb-3 flex items-center gap-2 rounded-md px-2 py-2 text-[13px] hover:bg-white/15'
				>
					<ArrowLeft size={16} /> {t('file.back')}
				</button>
				{shell.newWorkbook && (
					<button
						type='button'
						onClick={run(shell.newWorkbook)}
						className='flex items-center gap-2 rounded-md px-2 py-2 text-[13px] hover:bg-white/15'
					>
						<FilePlus2 size={15} /> {t('file.new')}
					</button>
				)}
				{shell.openFile && (
					<button
						type='button'
						onClick={run(shell.openFile)}
						className='flex items-center gap-2 rounded-md px-2 py-2 text-[13px] hover:bg-white/15'
					>
						<FolderOpen size={15} /> {t('file.open')}
					</button>
				)}
				<button
					type='button'
					disabled={store.readOnly}
					onClick={run(shell.save)}
					className='flex items-center gap-2 rounded-md px-2 py-2 text-[13px] hover:bg-white/15 disabled:opacity-50'
				>
					<Save size={15} /> {t('file.save')}
				</button>
				<button
					type='button'
					aria-pressed={!printing}
					onClick={() => setPrinting(false)}
					className={`flex items-center gap-2 rounded-md px-2 py-2 text-[13px] hover:bg-white/15 ${printing ? '' : 'bg-white/20'}`}
				>
					<Info size={15} /> {t('file.info')}
				</button>
				<button
					type='button'
					aria-pressed={printing}
					onClick={() => setPrinting(true)}
					className={`flex items-center gap-2 rounded-md px-2 py-2 text-[13px] hover:bg-white/15 ${printing ? 'bg-white/20' : ''}`}
				>
					<Printer size={15} /> {t('file.print')}
				</button>
			</nav>
			{printing ? (
				<PrintView fileName={fileName} />
			) : (
				<div className='flex min-w-0 flex-1 gap-8 overflow-y-auto p-8'>
					<section className='flex w-[420px] shrink-0 flex-col gap-2'>
						<h2 className='mb-2 text-[22px] font-semibold'>{t('file.export')}</h2>
						<Row
							icon={<FileSpreadsheet size={20} />}
							label={t('file.downloadXlsx')}
							detail={t('backstage.downloadXlsxDetail')}
							onClick={run(() => shell.download('xlsx'))}
						/>
						<Row
							icon={<FileText size={20} />}
							label={t('file.downloadCsv')}
							detail={t('backstage.downloadCsvDetail')}
							onClick={run(() => shell.download('csv'))}
						/>
						<Row
							icon={<FileText size={20} />}
							label={t('file.exportMarkdown')}
							detail={t('backstage.exportMarkdownDetail')}
							onClick={run(() => shell.download('markdown'))}
						/>
						<Row
							icon={<FileJson size={20} />}
							label={t('file.exportJson')}
							detail={t('backstage.exportJsonDetail')}
							onClick={run(() => shell.download('json'))}
						/>
						<Row
							icon={<Globe size={20} />}
							label={t('file.exportHtml')}
							detail={t('backstage.exportHtmlDetail')}
							onClick={run(() => shell.download('html'))}
						/>
						<Row
							icon={<Download size={20} />}
							label={t('file.print')}
							detail={t('backstage.printDetail')}
							onClick={() => setPrinting(true)}
						/>
					</section>
					<section className='flex min-w-[260px] flex-col gap-3'>
						<h2 className='mb-2 flex items-center gap-2 text-[22px] font-semibold'>
							<Info size={20} /> {t('file.info')}
						</h2>
						<ProtectWorkbook />
						<dl className='grid grid-cols-[120px_1fr] gap-x-4 gap-y-2 text-[13px]'>
							<dt className='text-muted-foreground'>{t('backstage.file')}</dt>
							<dd className='truncate'>{fileName}</dd>
							<dt className='text-muted-foreground'>{t('backstage.sheets')}</dt>
							<dd>{formatNumber(wb.sheets.length, locale)}</dd>
							<dt className='text-muted-foreground'>{t('backstage.cells')}</dt>
							<dd>{formatNumber(cellCount, locale)}</dd>
							{props.title && (
								<>
									<dt className='text-muted-foreground'>{t('backstage.title')}</dt>
									<dd>{props.title}</dd>
								</>
							)}
							{props.creator && (
								<>
									<dt className='text-muted-foreground'>{t('backstage.author')}</dt>
									<dd>{props.creator}</dd>
								</>
							)}
							{props.lastModifiedBy && (
								<>
									<dt className='text-muted-foreground'>{t('backstage.lastModifiedBy')}</dt>
									<dd>{props.lastModifiedBy}</dd>
								</>
							)}
							{props.modified && (
								<>
									<dt className='text-muted-foreground'>{t('backstage.modified')}</dt>
									<dd>{new Date(props.modified).toLocaleString(locale)}</dd>
								</>
							)}
							{props.application && (
								<>
									<dt className='text-muted-foreground'>{t('backstage.application')}</dt>
									<dd>{props.application}</dd>
								</>
							)}
						</dl>
						{wb.warnings.length > 0 && (
							<div className='mt-4 rounded-md border border-border bg-card p-3'>
								<h3 className='mb-2 text-[13px] font-semibold'>
									{t('backstage.compatibilityNotes')}
								</h3>
								<ul className='list-disc space-y-1 pl-5 text-[12.5px] text-muted-foreground'>
									{wb.warnings.slice(0, 12).map((w, i) => (
										// biome-ignore lint: warnings are static
										<li key={i}>{w.message}</li>
									))}
								</ul>
							</div>
						)}
					</section>
				</div>
			)}
		</div>
	);
}
