import { enChrome } from './locales/en/chrome';
import { enCommands } from './locales/en/commands';
import { enDialogs } from './locales/en/dialogs';
import { enDialogsData } from './locales/en/dialogs-data';
import { enDialogsRules } from './locales/en/dialogs-rules';
import { enDialogsTools } from './locales/en/dialogs-tools';
import { enFormatCells } from './locales/en/format-cells';
import { enPrint } from './locales/en/print';
import { enRibbon } from './locales/en/ribbon';
import { enRibbonCommands } from './locales/en/ribbon-commands';
import { enRibbonContextual } from './locales/en/ribbon-contextual';

/**
 * The English UI strings: the canonical key set. Each area lives in its own
 * file under `locales/en/`; every other built-in locale mirrors those files
 * key for key (see `locales/<code>/`).
 */
export const EN_AREAS = {
	chrome: enChrome,
	ribbon: enRibbon,
	ribbonCommands: enRibbonCommands,
	ribbonContextual: enRibbonContextual,
	dialogs: enDialogs,
	formatCells: enFormatCells,
	dialogsData: enDialogsData,
	dialogsRules: enDialogsRules,
	dialogsTools: enDialogsTools,
	print: enPrint,
	commands: enCommands,
} as const;

export const en = {
	...enChrome,
	...enRibbon,
	...enRibbonCommands,
	...enRibbonContextual,
	...enDialogs,
	...enFormatCells,
	...enDialogsData,
	...enDialogsRules,
	...enDialogsTools,
	...enPrint,
	...enCommands,
} as const;

export type TranslationKey = keyof typeof en;

/** Partial UI string overrides, keyed like the English dictionary. */
export type Translations = Partial<Record<TranslationKey, string>>;
