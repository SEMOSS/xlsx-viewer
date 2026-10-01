import type { TranslationDictionary } from '../types';
import { zhCNChrome } from './chrome';
import { zhCNCommands } from './commands';
import { zhCNDialogsData } from './dialogs-data';
import { zhCNDialogsRules } from './dialogs-rules';
import { zhCNDialogsTools } from './dialogs-tools';
import { zhCNDialogs } from './dialogs';
import { zhCNFormatCells } from './format-cells';
import { zhCNPrint } from './print';
import { zhCNRibbonCommands } from './ribbon-commands';
import { zhCNRibbonContextual } from './ribbon-contextual';
import { zhCNRibbon } from './ribbon';

/** Simplified Chinese UI strings, in Microsoft Office terminology. */
export const zhCN: TranslationDictionary = {
	...zhCNChrome,
	...zhCNCommands,
	...zhCNDialogsData,
	...zhCNDialogsRules,
	...zhCNDialogsTools,
	...zhCNDialogs,
	...zhCNFormatCells,
	...zhCNPrint,
	...zhCNRibbonCommands,
	...zhCNRibbonContextual,
	...zhCNRibbon,
};
