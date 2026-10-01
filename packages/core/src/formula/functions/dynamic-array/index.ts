import type { FunctionDefinition } from '../../types';
import { FILTER } from './filter';
import { GROUPBY } from './groupby';
import { PIVOTBY } from './pivotby';
import { HSTACK, TOCOL, TOROW, VSTACK, WRAPCOLS, WRAPROWS } from './reshape';
import { RANDARRAY, SEQUENCE } from './sequence';
import { CHOOSECOLS, CHOOSEROWS, DROP, EXPAND, TAKE } from './slice';
import { SORT, SORTBY } from './sort';
import { UNIQUE } from './unique';

/** Dynamic-array functions (Excel 365): array-shaping, sorting, filtering and generating. */
export const DYNAMIC_ARRAY_FUNCTIONS: FunctionDefinition[] = [
	FILTER,
	SORT,
	SORTBY,
	UNIQUE,
	SEQUENCE,
	RANDARRAY,
	TAKE,
	DROP,
	CHOOSECOLS,
	CHOOSEROWS,
	EXPAND,
	VSTACK,
	HSTACK,
	TOCOL,
	TOROW,
	WRAPROWS,
	WRAPCOLS,
	GROUPBY,
	PIVOTBY,
];

export {
	CHOOSECOLS,
	CHOOSEROWS,
	DROP,
	EXPAND,
	FILTER,
	GROUPBY,
	HSTACK,
	PIVOTBY,
	RANDARRAY,
	SEQUENCE,
	SORT,
	SORTBY,
	TAKE,
	TOCOL,
	TOROW,
	UNIQUE,
	VSTACK,
	WRAPCOLS,
	WRAPROWS,
};
