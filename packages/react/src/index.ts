/**
 * xlsx-react-viewer: React viewer and editor for Excel workbooks.
 *
 * ```tsx
 * import { ExcelViewer } from 'xlsx-react-viewer';
 * import 'xlsx-react-viewer/styles';
 *
 * <ExcelViewer content={bytes} fileName='Budget.xlsx' onSave={(b) => upload(b)} />
 * ```
 *
 * @packageDocumentation
 */
export * from './viewer';
export * from './theme';
export {
	BUILT_IN_DICTIONARIES,
	BUILT_IN_LOCALES,
	createTranslator,
	de as germanTranslations,
	en as englishTranslations,
	es as spanishTranslations,
	fr as frenchTranslations,
	I18nProvider,
	isSupportedLocale,
	useLocale,
	useT,
	zhCN as simplifiedChineseTranslations,
} from './i18n';
export type {
	I18nProviderProps,
	LocaleInfo,
	SupportedLocale,
	Translate,
	TranslationDictionary,
	TranslationKey,
	Translations,
	TranslationVars,
} from './i18n';

/** File extensions the viewer opens (for `<input accept>`). */
export const XLSX_OPEN_ACCEPT =
	'.xlsx,.xlsm,.xltx,.xltm,.xls,.csv,.tsv,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet,application/vnd.ms-excel.sheet.macroEnabled.12,application/vnd.ms-excel,text/csv,text/tab-separated-values';

/** The file name looks like a workbook the viewer can open. */
export function isSupportedWorkbookFile(fileName: string): boolean {
	return /\.(xlsx|xlsm|xltx|xltm|xls|csv|tsv)$/i.test(fileName);
}
