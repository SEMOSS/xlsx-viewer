import type { TranslationDictionary } from '../types';
import { deChrome } from './chrome';
import { deCommands } from './commands';
import { deDialogsData } from './dialogs-data';
import { deDialogsRules } from './dialogs-rules';
import { deDialogsTools } from './dialogs-tools';
import { deDialogs } from './dialogs';
import { deFormatCells } from './format-cells';
import { dePrint } from './print';
import { deRibbonCommands } from './ribbon-commands';
import { deRibbonContextual } from './ribbon-contextual';
import { deRibbon } from './ribbon';

/** German UI strings, in Microsoft Office terminology. */
export const de: TranslationDictionary = {
	...deChrome,
	...deCommands,
	...deDialogsData,
	...deDialogsRules,
	...deDialogsTools,
	...deDialogs,
	...deFormatCells,
	...dePrint,
	...deRibbonCommands,
	...deRibbonContextual,
	...deRibbon,
};
