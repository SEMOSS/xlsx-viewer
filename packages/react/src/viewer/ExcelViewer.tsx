import type { Workbook, WriteOptions } from 'xlsx-viewer-core';
import { createWorkbook, readWorkbook, writeWorkbook, XlsxPasswordError } from 'xlsx-viewer-core';
import {
	forwardRef,
	useCallback,
	useEffect,
	useImperativeHandle,
	useMemo,
	useRef,
	useState,
	type CSSProperties,
} from 'react';

import { createTranslator, I18nProvider } from '../i18n';
import { jadeDarkTheme, jadeLightTheme, themeToCssVars, type ViewerTheme } from '../theme';
import { goToReference } from './commands/selection-commands';
import { PasswordDialog } from './components/chrome/PasswordDialog';
import {
	WorkbookPasswordContext,
	type WorkbookPasswordState,
} from './components/chrome/password-context';
import { DialogButton } from './components/ui/Dialog';
import { downloadBlob, savedFileName, setWorkbookPassword } from './download';
import { ViewerShell } from './ViewerShell';
import { ViewerStore } from './state/viewer-store';
import { ViewerStoreContext } from './state/use-store';
import type { ExcelViewerHandle, ExcelViewerProps } from './types';

function resolveThemeProp(theme: ExcelViewerProps['theme'], mode: 'light' | 'dark'): ViewerTheme {
	if (theme && typeof theme === 'object') {
		return theme;
	}
	return mode === 'dark' ? jadeDarkTheme : jadeLightTheme;
}

type LoadState =
	| { status: 'loading' }
	| { status: 'error'; error: Error }
	/** The file is encrypted: ask for the password (again, when the last one was wrong). */
	| { status: 'password'; incorrect: boolean }
	| { status: 'ready'; store: ViewerStore };

/** `required` / `incorrect` when reading failed for want of the right password. */
function passwordProblem(e: unknown): 'required' | 'incorrect' | undefined {
	if (!(e instanceof Error) || e.name !== 'XlsxPasswordError') {
		return undefined;
	}
	const { reason } = e as XlsxPasswordError;
	return reason === 'required' || reason === 'incorrect' ? reason : undefined;
}

/** Retry opening the content with a typed password, or give up. */
interface Unlock {
	open: (password: string) => void;
	cancel: () => void;
}

const NO_UNLOCK: Unlock = { open: () => undefined, cancel: () => undefined };

/** Saves are encrypted with the workbook password, when there is one. */
function writeOptions(password: string | undefined): WriteOptions {
	return password ? { password } : {};
}

/**
 * Excel viewer and editor. Pass workbook bytes as `content` (or a parsed
 * `workbook`); edits stay in memory until saved through `onSave` / the
 * imperative handle.
 *
 * ```tsx
 * <ExcelViewer content={bytes} fileName='Budget.xlsx' onSave={(b) => upload(b)} />
 * ```
 */
export const ExcelViewer = forwardRef<ExcelViewerHandle, ExcelViewerProps>(
	function ExcelViewer(props, ref) {
		const { content, workbook, canEdit = true } = props;
		const [load, setLoad] = useState<LoadState>({ status: 'loading' });
		const initialMode =
			props.theme === 'dark' || (typeof props.theme === 'object' && props.theme.mode === 'dark')
				? 'dark'
				: 'light';
		const [mode, setMode] = useState<'light' | 'dark'>(initialMode);
		const locale = props.locale ?? 'en';
		const t = useMemo(
			() => createTranslator(props.translations, locale),
			[props.translations, locale],
		);
		// The store translates through this ref, so a new locale applies without reloading.
		const tRef = useRef(t);
		tRef.current = t;
		const propsRef = useRef(props);
		propsRef.current = props;
		// The password saves are encrypted with: viewer state, never stored in the workbook.
		const [password, setPassword] = useState<string | undefined>();
		const passwordRef = useRef<string | undefined>(undefined);
		const unlockRef = useRef<Unlock>(NO_UNLOCK);

		useEffect(() => {
			if (props.theme === 'light' || props.theme === 'dark') {
				setMode(props.theme);
			} else if (props.theme && typeof props.theme === 'object' && props.theme.mode) {
				setMode(props.theme.mode);
			}
		}, [props.theme]);

		// Load (or create) the workbook whenever the input changes.
		useEffect(() => {
			let cancelled = false;
			setLoad({ status: 'loading' });
			const done = (wb: Workbook, openPassword: string | undefined) => {
				if (!cancelled) {
					// As in Excel, saving re-encrypts with the password the file was opened with.
					const kept = wb.encrypted ? openPassword || undefined : undefined;
					passwordRef.current = kept;
					setPassword(kept);
					const ready = new ViewerStore(wb, {
						readOnly: !canEdit,
						t: (key, vars) => tRef.current(key, vars),
					});
					setLoad({ status: 'ready', store: ready });
				}
			};
			const fail = (e: unknown) => {
				if (cancelled) {
					return;
				}
				const error = e instanceof Error ? e : new Error(String(e));
				setLoad({ status: 'error', error });
				propsRef.current.onError?.(error);
			};
			const open = (bytes: Uint8Array | ArrayBuffer, openPassword: string | undefined) => {
				// The file name lets the reader open .csv / .tsv text as a workbook.
				const fileName = propsRef.current.fileName;
				readWorkbook(
					bytes,
					openPassword === undefined ? { fileName } : { password: openPassword, fileName },
				)
					.then((wb) => done(wb, openPassword))
					.catch((e: unknown) => {
						const problem = passwordProblem(e);
						if (problem && !cancelled) {
							setLoad({ status: 'password', incorrect: problem === 'incorrect' });
						} else {
							fail(e);
						}
					});
			};
			if (workbook) {
				done(workbook, propsRef.current.password);
			} else if (content) {
				unlockRef.current = {
					open: (entered) => {
						setLoad({ status: 'loading' });
						open(content, entered);
					},
					cancel: () => {
						fail(new XlsxPasswordError(tRef.current('password.cancelled'), 'required'));
					},
				};
				open(content, propsRef.current.password);
			} else {
				done(createWorkbook({ title: propsRef.current.fileName }), undefined);
			}
			return () => {
				cancelled = true;
			};
			// canEdit is applied below without reloading; the password prop is read on load.
			// eslint-disable-next-line react-hooks/exhaustive-deps
		}, [content, workbook]);

		const store = load.status === 'ready' ? load.store : null;

		useEffect(() => {
			if (store && store.readOnly !== !canEdit) {
				store.readOnly = !canEdit;
				// Re-render only: switching edit mode does not change the document.
				store.setUi({});
			}
		}, [store, canEdit]);

		// `.xlsx` downloads (File > Export) look the password up by workbook.
		useEffect(() => {
			if (store) {
				setWorkbookPassword(store.workbook, password);
			}
		}, [store, password]);

		const passwordState: WorkbookPasswordState = useMemo(
			() => ({
				password,
				setPassword: (next) => {
					const value = next || undefined;
					passwordRef.current = value;
					setPassword(value);
					if (store) {
						setWorkbookPassword(store.workbook, value);
						// Adding or removing the password changes the saved file, like an edit.
						store.touch();
					}
				},
			}),
			[password, store],
		);

		// Wire host callbacks.
		useEffect(() => {
			if (!store) {
				return;
			}
			let timer: ReturnType<typeof setTimeout> | undefined;
			store.callbacks = {
				onDirtyChange: (d) => propsRef.current.onDirtyChange?.(d),
				onActiveSheetChange: (i) => propsRef.current.onActiveSheetChange?.(i),
				onSelectionChange: (sheet, sel) =>
					propsRef.current.onSelectionChange?.({ sheet, active: sel.active, ranges: sel.ranges }),
				onChange: () => {
					if (!propsRef.current.onContentChange) {
						return;
					}
					clearTimeout(timer);
					timer = setTimeout(() => {
						void writeWorkbook(store.workbook, writeOptions(passwordRef.current)).then((bytes) =>
							propsRef.current.onContentChange?.(bytes),
						);
					}, propsRef.current.contentChangeDelayMs ?? 800);
				},
			};
			return () => {
				clearTimeout(timer);
				store.callbacks = {};
			};
		}, [store]);

		const theme = useMemo(() => resolveThemeProp(props.theme, mode), [props.theme, mode]);
		const cssVars = useMemo(() => themeToCssVars({ ...theme, mode }), [theme, mode]);

		const toggleTheme = useCallback(() => {
			setMode((m) => {
				const next = m === 'dark' ? 'light' : 'dark';
				propsRef.current.onThemeModeChange?.(next);
				return next;
			});
		}, []);

		const save = useCallback(async () => {
			if (!store) {
				return;
			}
			const bytes = await writeWorkbook(store.workbook, writeOptions(passwordRef.current));
			const name = savedFileName(propsRef.current.fileName ?? 'Workbook.xlsx');
			if (propsRef.current.onSave) {
				await propsRef.current.onSave(bytes, name);
			} else {
				downloadBlob(
					bytes as BlobPart,
					name,
					'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
				);
			}
			store.markSaved();
			propsRef.current.onContentChange?.(bytes);
		}, [store]);

		const focusRef = useRef<() => void>(() => undefined);

		useImperativeHandle(
			ref,
			(): ExcelViewerHandle => ({
				getContent: async () => {
					if (!store) {
						throw new Error('Workbook not loaded');
					}
					return writeWorkbook(store.workbook, writeOptions(passwordRef.current));
				},
				getWorkbook: () => store?.workbook ?? null,
				save,
				undo: () => store?.undo(),
				redo: () => store?.redo(),
				recalculate: () => store?.recalculate(),
				select: (reference) => (store ? goToReference(store, reference) : false),
				setActiveSheet: (index) => store?.setUi({ activeSheet: index }),
				focus: () => focusRef.current(),
			}),
			[store, save],
		);

		const rootStyle: CSSProperties = {
			...(cssVars as CSSProperties),
			colorScheme: mode,
			...props.style,
		};

		return (
			<I18nProvider locale={locale} translations={props.translations}>
				<div
					data-xlsx-viewer=''
					data-theme={mode}
					className={`xv-root ${mode === 'dark' ? 'dark' : ''} ${props.className ?? ''}`}
					style={rootStyle}
				>
					{load.status === 'loading' && (
						<div className='flex h-full w-full items-center justify-center gap-3 text-[13px] text-muted-foreground'>
							<span className='xv-spinner' aria-hidden />
							{t('app.loading')}
						</div>
					)}
					{load.status === 'password' && (
						<PasswordDialog
							fileName={props.fileName ?? 'Workbook.xlsx'}
							incorrect={load.incorrect}
							onSubmit={(entered) => unlockRef.current.open(entered)}
							onCancel={() => unlockRef.current.cancel()}
						/>
					)}
					{load.status === 'error' && (
						<div
							role='alert'
							className='flex h-full w-full flex-col items-center justify-center gap-2 p-8 text-center'
						>
							<h2 className='text-[16px] font-semibold text-foreground'>{t('app.error.title')}</h2>
							<p className='max-w-[520px] text-[13px] text-muted-foreground'>
								{load.error.message}
							</p>
							{passwordProblem(load.error) && content && (
								<DialogButton onClick={() => setLoad({ status: 'password', incorrect: false })}>
									{t('password.enter')}
								</DialogButton>
							)}
						</div>
					)}
					{store && (
						<ViewerStoreContext.Provider value={store}>
							<WorkbookPasswordContext.Provider value={passwordState}>
								<ViewerShell
									fileName={props.fileName ?? 'Book1.xlsx'}
									canEdit={canEdit}
									mode={mode}
									onToggleTheme={toggleTheme}
									onSave={save}
									onOpenFile={props.onOpenFile}
									onNewWorkbook={props.onNewWorkbook}
									authorName={props.authorName ?? 'Author'}
									showRibbon={props.showRibbon ?? true}
									showFormulaBar={props.showFormulaBar ?? true}
									showSheetTabs={props.showSheetTabs ?? true}
									showStatusBar={props.showStatusBar ?? true}
									showTitleBar={props.showTitleBar ?? true}
									registerFocus={(fn) => {
										focusRef.current = fn;
									}}
								/>
							</WorkbookPasswordContext.Provider>
						</ViewerStoreContext.Provider>
					)}
				</div>
			</I18nProvider>
		);
	},
);
