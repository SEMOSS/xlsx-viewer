import type { TranslationDictionary } from '../types';
import { frChrome } from './chrome';
import { frCommands } from './commands';
import { frDialogsData } from './dialogs-data';
import { frDialogsRules } from './dialogs-rules';
import { frDialogsTools } from './dialogs-tools';
import { frDialogs } from './dialogs';
import { frFormatCells } from './format-cells';
import { frPrint } from './print';
import { frRibbonCommands } from './ribbon-commands';
import { frRibbonContextual } from './ribbon-contextual';
import { frRibbon } from './ribbon';

/** French UI strings, in Microsoft Office terminology. */
export const fr: TranslationDictionary = {
	...frChrome,
	...frCommands,
	...frDialogsData,
	...frDialogsRules,
	...frDialogsTools,
	...frDialogs,
	...frFormatCells,
	...frPrint,
	...frRibbonCommands,
	...frRibbonContextual,
	...frRibbon,
};
