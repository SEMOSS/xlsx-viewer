import { useCallback, useEffect, useMemo, useRef, useState, type DragEvent } from 'react';
import { createRoot } from 'react-dom/client';

import type { Workbook } from 'xlsx-viewer-core';
import { createWorkbook } from 'xlsx-viewer-core';
import {
	ExcelViewer,
	isSupportedWorkbookFile,
	themeToCssVars,
	XLSX_OPEN_ACCEPT,
	type ExcelViewerHandle,
	type SupportedLocale,
} from 'xlsx-react-viewer';

import { LanguagePicker } from './LanguagePicker';
import { readLocale, saveLocale } from './locale';
import { buildSampleWorkbook } from './sample';
import { buildDynamicArraysWorkbook } from './samples/dynamic';
import { buildLoanWorkbook } from './samples/loan';
import { buildProjectWorkbook } from './samples/projects';
import './app.css';

type Doc =
	| { kind: 'bytes'; bytes: Uint8Array; name: string }
	| { kind: 'workbook'; workbook: Workbook; name: string };

/** Sample workbooks built in code, each showing a different part of the viewer. */
const SAMPLES: Array<{
	id: string;
	name: string;
	title: string;
	detail: string;
	build: () => Workbook;
}> = [
	{
		id: 'sales',
		name: 'Quarterly Sales.xlsx',
		title: 'Quarterly sales',
		detail:
			'Formulas, accounting formats, data bars, color scales, icons, sparklines, a chart and a table.',
		build: buildSampleWorkbook,
	},
	{
		id: 'loan',
		name: 'Loan Calculator.xlsx',
		title: 'Loan calculator',
		detail:
			'PMT, CUMIPMT and EDATE, validated input cells, a 30-year schedule and a balance chart.',
		build: buildLoanWorkbook,
	},
	{
		id: 'projects',
		name: 'Project Tracker.xlsx',
		title: 'Project tracker',
		detail: 'A styled table with status lists, highlight rules, progress icons and grouped phases.',
		build: buildProjectWorkbook,
	},
	{
		id: 'dynamic',
		name: 'Dynamic Arrays.xlsx',
		title: 'Dynamic arrays',
		detail: 'FILTER, SORT, UNIQUE, SEQUENCE, XLOOKUP, LET and LAMBDA, with spilled results.',
		build: buildDynamicArraysWorkbook,
	},
];

function readTheme(): 'light' | 'dark' {
	try {
		const stored = localStorage.getItem('xlsx-demo-theme');
		if (stored === 'light' || stored === 'dark') {
			return stored;
		}
	} catch {
		/* ignore */
	}
	return window.matchMedia?.('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
}

function App() {
	const [doc, setDoc] = useState<Doc | null>(null);
	const [error, setError] = useState<string | null>(null);
	const [over, setOver] = useState(false);
	const [mode, setMode] = useState<'light' | 'dark'>(readTheme);
	const [locale, setLocale] = useState<SupportedLocale>(readLocale);
	const inputRef = useRef<HTMLInputElement>(null);
	const viewerRef = useRef<ExcelViewerHandle>(null);

	useEffect(() => {
		try {
			localStorage.setItem('xlsx-demo-theme', mode);
		} catch {
			/* ignore */
		}
		document.documentElement.style.colorScheme = mode;
	}, [mode]);

	useEffect(() => {
		saveLocale(locale);
		document.documentElement.lang = locale;
	}, [locale]);

	const vars = useMemo(() => themeToCssVars({ mode }), [mode]);

	const openFile = useCallback(async (file: File) => {
		setError(null);
		if (!isSupportedWorkbookFile(file.name)) {
			setError(
				`${file.name} is not a workbook the viewer can open (.xlsx, .xlsm, .xls, .csv or .tsv).`,
			);
			return;
		}
		const bytes = new Uint8Array(await file.arrayBuffer());
		setDoc({ kind: 'bytes', bytes, name: file.name });
		document.title = `${file.name} - XLSX Viewer`;
	}, []);

	const onDrop = useCallback(
		(e: DragEvent) => {
			e.preventDefault();
			setOver(false);
			const file = e.dataTransfer.files[0];
			if (file) {
				void openFile(file);
			}
		},
		[openFile],
	);

	const input = (
		<input
			ref={inputRef}
			type='file'
			id='file-input'
			accept={XLSX_OPEN_ACCEPT}
			className='sr-only'
			aria-label='Open a workbook'
			onChange={(e) => {
				const file = e.target.files?.[0];
				if (file) {
					void openFile(file);
				}
				e.target.value = '';
			}}
		/>
	);

	if (doc) {
		return (
			<main style={{ height: '100dvh', width: '100vw' }}>
				<ExcelViewer
					ref={viewerRef}
					content={doc.kind === 'bytes' ? doc.bytes : undefined}
					workbook={doc.kind === 'workbook' ? doc.workbook : undefined}
					fileName={doc.name}
					locale={locale}
					theme={mode}
					onThemeModeChange={setMode}
					onOpenFile={() => inputRef.current?.click()}
					onNewWorkbook={() =>
						setDoc({ kind: 'workbook', workbook: buildBlank(), name: 'Book1.xlsx' })
					}
					onDirtyChange={(dirty) => {
						document.title = `${dirty ? '* ' : ''}${doc.name} - XLSX Viewer`;
					}}
					authorName='Demo User'
				/>
				<LanguagePicker
					value={locale}
					onChange={setLocale}
					className='demo-locale-floating'
					style={vars as React.CSSProperties}
				/>
				{input}
			</main>
		);
	}

	return (
		<main className='demo-stage' style={vars as React.CSSProperties}>
			<div className='demo-top-actions'>
				<LanguagePicker value={locale} onChange={setLocale} />
				<button
					type='button'
					className='demo-theme-toggle'
					onClick={() => setMode(mode === 'dark' ? 'light' : 'dark')}
				>
					{mode === 'dark' ? 'Light mode' : 'Dark mode'}
				</button>
			</div>
			<div className='demo-brand'>
				<img src='/favicon.svg' width={34} height={34} alt='' />
				XLSX Viewer
			</div>
			{/* biome-ignore lint: drag and drop supplements the keyboard-accessible file input */}
			<div
				className='demo-dropzone'
				data-testid='dropzone'
				data-over={over}
				role='group'
				aria-label='Open a workbook'
				onClick={(e) => {
					if (!(e.target as HTMLElement).closest('button, label, input')) {
						inputRef.current?.click();
					}
				}}
				onDragOver={(e) => {
					e.preventDefault();
					setOver(true);
				}}
				onDragLeave={() => setOver(false)}
				onDrop={onDrop}
			>
				<label className='demo-hint' htmlFor='file-input'>
					Drop an Excel workbook here, or click to browse
				</label>
				<p className='demo-sub'>
					.xlsx, .xlsm, legacy .xls and .csv files are opened entirely in your browser. Nothing is
					uploaded.
				</p>
				<div className='demo-actions'>
					<button
						type='button'
						className='demo-primary'
						data-testid='browse-files'
						onClick={() => inputRef.current?.click()}
					>
						Browse files
					</button>
					<button
						type='button'
						onClick={() => setDoc({ kind: 'workbook', workbook: buildBlank(), name: 'Book1.xlsx' })}
					>
						New workbook
					</button>
					<button
						type='button'
						data-testid='open-sample'
						onClick={() =>
							setDoc({
								kind: 'workbook',
								workbook: buildSampleWorkbook(),
								name: 'Quarterly Sales.xlsx',
							})
						}
					>
						Open sample
					</button>
				</div>
				{input}
				{error && <p className='demo-error'>{error}</p>}
			</div>
			<section className='demo-samples' aria-label='Sample workbooks'>
				<h2>Or open a sample</h2>
				<div className='demo-sample-grid'>
					{SAMPLES.map((sample) => (
						<button
							key={sample.id}
							type='button'
							className='demo-sample'
							data-testid={`sample-${sample.id}`}
							onClick={() => {
								setDoc({ kind: 'workbook', workbook: sample.build(), name: sample.name });
								document.title = `${sample.name} - XLSX Viewer`;
							}}
						>
							<strong>{sample.title}</strong>
							<span>{sample.detail}</span>
						</button>
					))}
				</div>
			</section>
			<p className='demo-footer'>React viewer and editor for Excel workbooks.</p>
		</main>
	);
}

function buildBlank(): Workbook {
	return createWorkbook();
}

const rootEl = document.getElementById('app-root');
if (rootEl) {
	const hot = import.meta.hot;
	const existing = hot?.data.root as ReturnType<typeof createRoot> | undefined;
	const root = existing ?? createRoot(rootEl);
	if (hot) {
		hot.data.root = root;
	}
	root.render(<App />);
}
