import type { TranslationDictionary } from '../types';
import { esChrome } from './chrome';
import { esCommands } from './commands';
import { esDialogsData } from './dialogs-data';
import { esDialogsRules } from './dialogs-rules';
import { esDialogsTools } from './dialogs-tools';
import { esDialogs } from './dialogs';
import { esFormatCells } from './format-cells';
import { esPrint } from './print';
import { esRibbonCommands } from './ribbon-commands';
import { esRibbonContextual } from './ribbon-contextual';
import { esRibbon } from './ribbon';

/** Spanish UI strings, in Microsoft Office terminology. */
export const es: TranslationDictionary = {
	...esChrome,
	...esCommands,
	...esDialogsData,
	...esDialogsRules,
	...esDialogsTools,
	...esDialogs,
	...esFormatCells,
	...esPrint,
	...esRibbonCommands,
	...esRibbonContextual,
	...esRibbon,
};
