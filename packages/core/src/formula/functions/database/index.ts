import type { FunctionDefinition } from '../../types';
import {
	DAVERAGE,
	DCOUNT,
	DCOUNTA,
	DGET,
	DMAX,
	DMIN,
	DPRODUCT,
	DSTDEV,
	DSTDEVP,
	DSUM,
	DVAR,
	DVARP,
} from './functions';

/** Database functions: aggregates over the records of a range that match a criteria range. */
export const DATABASE_FUNCTIONS: FunctionDefinition[] = [
	DAVERAGE,
	DCOUNT,
	DCOUNTA,
	DGET,
	DMAX,
	DMIN,
	DPRODUCT,
	DSTDEV,
	DSTDEVP,
	DSUM,
	DVAR,
	DVARP,
];

export {
	DAVERAGE,
	DCOUNT,
	DCOUNTA,
	DGET,
	DMAX,
	DMIN,
	DPRODUCT,
	DSTDEV,
	DSTDEVP,
	DSUM,
	DVAR,
	DVARP,
};
