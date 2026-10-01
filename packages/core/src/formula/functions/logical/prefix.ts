import { filePrefixFor } from '../../language/function-prefixes';
import type { FunctionDefinition } from '../../types';

/** Attach the file-format prefix (`_xlfn.`, `_xlfn._xlws.`) a function needs, per `formula/prefixes`. */
export function withFilePrefix(def: FunctionDefinition): FunctionDefinition {
	const prefix = filePrefixFor(def.name);
	return prefix === '_xlfn.' || prefix === '_xlfn._xlws.' ? { ...def, filePrefix: prefix } : def;
}
