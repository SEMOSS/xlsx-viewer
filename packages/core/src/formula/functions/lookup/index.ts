import type { FunctionDefinition } from '../../types';
import { LOOKUP } from './lookup-vector';
import { MATCH } from './match';
import { GETPIVOTDATA, HYPERLINK, TRANSPOSE } from './misc';
import { HLOOKUP, VLOOKUP } from './vlookup';
import { XLOOKUP, XMATCH } from './xlookup';

/** Lookup functions: VLOOKUP, HLOOKUP, LOOKUP, MATCH, XLOOKUP, XMATCH, TRANSPOSE, HYPERLINK, GETPIVOTDATA. */
export const LOOKUP_FUNCTIONS: FunctionDefinition[] = [
	VLOOKUP,
	HLOOKUP,
	LOOKUP,
	MATCH,
	XLOOKUP,
	XMATCH,
	TRANSPOSE,
	HYPERLINK,
	GETPIVOTDATA,
];

export { GETPIVOTDATA, HLOOKUP, HYPERLINK, LOOKUP, MATCH, TRANSPOSE, VLOOKUP, XLOOKUP, XMATCH };
