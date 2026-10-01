import type { FunctionDefinition } from '../../types';
import { ACCRUED_FUNCTIONS } from './accrued';
import { AMORTIZATION_FUNCTIONS } from './amortization';
import { BOND_FUNCTIONS } from './bonds';
import { CASHFLOW_FUNCTIONS } from './cashflow';
import { COUPON_FUNCTIONS } from './coupons';
import { DEPRECIATION_FUNCTIONS } from './depreciation';
import { DISCOUNT_FUNCTIONS } from './discount';
import { ODD_COUPON_FUNCTIONS } from './odd-coupon';
import { RATE_FUNCTIONS } from './rate';
import { TBILL_FUNCTIONS } from './tbill';
import { TVM_FUNCTIONS } from './tvm';

/** Financial functions (PMT, RATE, XIRR, PRICE, ODDFPRICE, VDB, ...). */
export const FINANCIAL_FUNCTIONS: FunctionDefinition[] = [
	...TVM_FUNCTIONS,
	...RATE_FUNCTIONS,
	...CASHFLOW_FUNCTIONS,
	...DEPRECIATION_FUNCTIONS,
	...AMORTIZATION_FUNCTIONS,
	...COUPON_FUNCTIONS,
	...BOND_FUNCTIONS,
	...ACCRUED_FUNCTIONS,
	...DISCOUNT_FUNCTIONS,
	...TBILL_FUNCTIONS,
	...ODD_COUPON_FUNCTIONS,
];
