/**
 * Treasury bills (TBILLEQ, TBILLPRICE, TBILLYIELD) and fractional dollar
 * prices (DOLLARDE, DOLLARFR).
 *
 * T-bills count actual days on a 360-day year and may not mature more than
 * one year after settlement.
 */
import type { CellError } from '../../../types/cell';
import { ERRORS, isCellError } from '../../../values/errors';
import type { FunctionArg, FunctionContext, FunctionDefinition } from '../../types';
import { finite, numArg, readDates, readNumbers, REQUIRED } from '../datetime/args';
import { addMonths, serialToYmd } from '../datetime/serial';

interface Bill {
	/** Days from settlement to maturity. */
	dsm: number;
	value: number;
}

function readBill(args: FunctionArg[], ctx: FunctionContext): Bill | CellError {
	const dates = readDates(args, ctx, 2, ERRORS.VALUE);
	const value = numArg(args[2], ctx);
	if (isCellError(dates)) {
		return dates;
	}
	if (isCellError(value)) {
		return value;
	}
	const [settlement, maturity] = dates;
	const oneYear = addMonths(serialToYmd(settlement, ctx.date1904), 12, false, ctx.date1904);
	if (settlement >= maturity || maturity > oneYear || value <= 0) {
		return ERRORS.NUM;
	}
	return { dsm: maturity - settlement, value };
}

function tbillEqImpl(args: FunctionArg[], ctx: FunctionContext) {
	const b = readBill(args, ctx);
	if (isCellError(b)) {
		return b;
	}
	const denominator = 360 - b.value * b.dsm;
	return denominator <= 0 ? ERRORS.NUM : finite((365 * b.value) / denominator);
}

function tbillPriceImpl(args: FunctionArg[], ctx: FunctionContext) {
	const b = readBill(args, ctx);
	if (isCellError(b)) {
		return b;
	}
	const price = 100 * (1 - (b.value * b.dsm) / 360);
	return price <= 0 ? ERRORS.NUM : price;
}

function tbillYieldImpl(args: FunctionArg[], ctx: FunctionContext) {
	const b = readBill(args, ctx);
	if (isCellError(b)) {
		return b;
	}
	return finite(((100 - b.value) / b.value) * (360 / b.dsm));
}

/** Smallest power of ten that is at least `fraction` (10^ceil(log10(fraction))). */
function decimalScale(fraction: number): number {
	let scale = 1;
	while (scale < fraction) {
		scale *= 10;
	}
	return scale;
}

/**
 * DOLLARDE / DOLLARFR: the digits after the decimal point are read (or
 * written) as a numerator over `fraction`. The fraction is truncated;
 * negative is #NUM!, below 1 is #DIV/0!.
 */
function dollar(toDecimal: boolean) {
	return (args: FunctionArg[], ctx: FunctionContext) => {
		const n = readNumbers(args, ctx, [REQUIRED, REQUIRED]);
		if (isCellError(n)) {
			return n;
		}
		const value = n[0];
		const fraction = Math.trunc(n[1]);
		if (fraction < 0) {
			return ERRORS.NUM;
		}
		if (fraction === 0) {
			return ERRORS.DIV0;
		}
		const whole = Math.trunc(value);
		const part = value - whole;
		const scale = decimalScale(fraction);
		return finite(whole + (toDecimal ? (part * scale) / fraction : (part * fraction) / scale));
	};
}

const def = (name: string, impl: FunctionDefinition['impl']): FunctionDefinition => ({
	name,
	minArgs: name.startsWith('DOLLAR') ? 2 : 3,
	maxArgs: name.startsWith('DOLLAR') ? 2 : 3,
	params: ['value'],
	impl,
});

export const TBILL_FUNCTIONS: FunctionDefinition[] = [
	def('TBILLEQ', tbillEqImpl),
	def('TBILLPRICE', tbillPriceImpl),
	def('TBILLYIELD', tbillYieldImpl),
	def('DOLLARDE', dollar(true)),
	def('DOLLARFR', dollar(false)),
];
