/**
 * Financial functions, part 1: time value of money, cash flows,
 * depreciation and rate conversions. Securities live in `financial-securities.ts`.
 */
import { arg, categoryBuilder, more, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Financial');

const RATE = arg('rate', 'Interest rate per period.');
const NPER = arg('nper', 'Total number of payment periods.');
const PMT = arg('pmt', 'Payment made each period; it stays the same over the whole term.');
const PV = arg('pv', 'Present value: the lump sum the future payments are worth now.');
const PV_OPT = opt(
	'pv',
	'Present value: the lump sum the future payments are worth now; 0 if omitted.',
);
const FV_OPT = opt('fv', 'Cash balance wanted after the last payment; 0 if omitted.');
const TYPE_OPT = opt(
	'type',
	'When payments are due: 0 or omitted at period end, 1 at period start.',
);
const PER = arg('per', 'The period to calculate, from 1 to nper.');
const GUESS = opt('guess', 'A starting estimate for the result; 10% if omitted.');
const COST = arg('cost', 'Initial cost of the asset.');
const SALVAGE = arg('salvage', 'Value of the asset at the end of its useful life.');
const LIFE = arg('life', 'Number of periods over which the asset is depreciated.');
const FACTOR = opt(
	'factor',
	'Rate at which the balance declines; 2 (double-declining) if omitted.',
);
const FRACTION = arg('fraction', 'The integer to use as the denominator of the fraction.');
const NPERY = arg('npery', 'Number of compounding periods per year.');

const CUMULATIVE_ARGS = [
	RATE,
	NPER,
	arg('pv', 'Present value: the amount of the loan.'),
	arg('start_period', 'First period in the calculation; periods are numbered from 1.'),
	arg('end_period', 'Last period in the calculation.'),
	arg('type', 'When payments are due: 0 at period end, 1 at period start.'),
];

const AMORTIZATION_ARGS = [
	COST,
	arg('date_purchased', 'Date the asset was purchased.'),
	arg('first_period', 'Date the first accounting period ends.'),
	arg('salvage', 'Salvage value at the end of the asset life.'),
	arg('period', 'The accounting period to calculate.'),
	arg('rate', 'Rate of depreciation.'),
	opt(
		'basis',
		'Year basis: 0 or omitted 360 days (NASD), 1 actual, 3 365 days, 4 360 days (European).',
	),
];

export const FINANCIAL_INFO: readonly FunctionInfo[] = [
	fn(
		'AMORDEGRC',
		'Calculates declining (degressive) depreciation for each accounting period (French system).',
		AMORTIZATION_ARGS,
	),
	fn(
		'AMORLINC',
		'Calculates prorated straight-line depreciation for each accounting period (French system).',
		AMORTIZATION_ARGS,
	),
	fn(
		'CUMIPMT',
		'Adds up the interest paid on a loan from one payment period to another.',
		CUMULATIVE_ARGS,
	),
	fn(
		'CUMPRINC',
		'Adds up the principal repaid on a loan from one payment period to another.',
		CUMULATIVE_ARGS,
	),
	fn('DB', 'Calculates fixed-declining balance depreciation of an asset for one period.', [
		COST,
		SALVAGE,
		LIFE,
		arg('period', 'The period to calculate, in the same units as life.'),
		opt('month', 'Number of months in the first year; 12 if omitted.'),
	]),
	fn('DDB', 'Calculates double-declining balance (or other factor) depreciation for one period.', [
		COST,
		SALVAGE,
		LIFE,
		arg('period', 'The period to calculate, in the same units as life.'),
		FACTOR,
	]),
	fn(
		'DOLLARDE',
		'Converts a dollar price written as a fraction, such as 1.02 for 1 2/16, to a decimal.',
		[arg('fractional_dollar', 'Price written as an integer part and a fraction part.'), FRACTION],
	),
	fn(
		'DOLLARFR',
		'Converts a decimal dollar price to one written as a fraction, such as 1.125 to 1.02.',
		[arg('decimal_dollar', 'The decimal number to convert.'), FRACTION],
	),
	fn(
		'EFFECT',
		'Converts a nominal annual rate into the effective rate, given the compounding frequency.',
		[arg('nominal_rate', 'The nominal annual interest rate.'), NPERY],
	),
	fn(
		'FV',
		'Calculates what an investment will be worth after a number of equal payments at a fixed rate.',
		[RATE, NPER, PMT, PV_OPT, TYPE_OPT],
	),
	fn(
		'FVSCHEDULE',
		'Grows a starting principal through a list of interest rates applied one after another.',
		[
			arg('principal', 'The present value.'),
			arg('schedule', 'An array or range of interest rates to apply in turn.'),
		],
	),
	fn(
		'IPMT',
		'Gives the interest portion of one payment of a loan with equal payments at a fixed rate.',
		[RATE, PER, NPER, PV, FV_OPT, TYPE_OPT],
	),
	fn(
		'IRR',
		'Finds the discount rate at which regular cash flows have a net present value of zero.',
		[
			arg('values', 'Cash flows; at least one payment (negative) and one receipt (positive).'),
			GUESS,
		],
	),
	fn(
		'ISPMT',
		'Returns the interest paid in a given period of a loan that repays equal amounts of principal.',
		[
			RATE,
			arg('per', 'The period to calculate; ISPMT counts periods from zero.'),
			NPER,
			arg('pv', 'Present value: the amount of the loan.'),
		],
	),
	fn(
		'MIRR',
		'Returns the modified internal rate of return, with separate financing and reinvestment rates.',
		[
			arg('values', 'Regular cash flows, with at least one negative and one positive value.'),
			arg('finance_rate', 'Interest rate paid on the money used in the cash flows.'),
			arg('reinvest_rate', 'Interest rate earned on cash flows when they are reinvested.'),
		],
	),
	fn(
		'NOMINAL',
		'Converts an effective annual rate into the nominal rate, given the compounding frequency.',
		[arg('effect_rate', 'The effective interest rate.'), NPERY],
	),
	fn(
		'NPER',
		'Counts the equal payments at a fixed rate needed to repay a loan or reach a savings goal.',
		[RATE, PMT, PV, FV_OPT, TYPE_OPT],
	),
	fn(
		'NPV',
		'Discounts a series of regular future cash flows to today at a fixed rate and adds them up.',
		[
			arg('rate', 'Discount rate for one period.'),
			arg('value1', 'The first cash flow, at the end of the first period.'),
			more('value2', 'More cash flows, equally spaced in time (up to 254 in total).'),
		],
	),
	fn(
		'PDURATION',
		'Calculates how many periods an investment at a fixed rate takes to grow to a target value.',
		[
			arg('rate', 'Interest rate per period.'),
			arg('pv', 'Present value of the investment.'),
			arg('fv', 'Desired future value of the investment.'),
		],
	),
	fn('PMT', 'Calculates the equal periodic payment that repays a loan at a fixed interest rate.', [
		RATE,
		NPER,
		arg('pv', 'Present value: the principal, or total the future payments are worth now.'),
		FV_OPT,
		TYPE_OPT,
	]),
	fn(
		'PPMT',
		'Gives the principal portion of one payment of a loan with equal payments at a fixed rate.',
		[RATE, PER, NPER, PV, FV_OPT, TYPE_OPT],
	),
	fn('PV', 'Calculates what a series of equal future payments at a fixed rate is worth today.', [
		RATE,
		NPER,
		PMT,
		FV_OPT,
		TYPE_OPT,
	]),
	fn('RATE', 'Solves for the interest rate per period of a loan or annuity with equal payments.', [
		NPER,
		PMT,
		PV,
		FV_OPT,
		TYPE_OPT,
		GUESS,
	]),
	fn(
		'RRI',
		'Calculates the per-period rate at which an investment grows from a present to a future value.',
		[
			arg('nper', 'Number of periods.'),
			arg('pv', 'Present value of the investment.'),
			arg('fv', 'Future value of the investment.'),
		],
	),
	fn(
		'SLN',
		'Calculates the equal per-period depreciation of an asset under the straight-line method.',
		[COST, SALVAGE, LIFE],
	),
	fn(
		'STOCKHISTORY',
		'Downloads historical prices of a stock or other instrument as an array (online service).',
		[
			arg('stock', 'Ticker symbol in quotes, or a cell holding a Stocks data type.'),
			arg('start_date', 'Earliest date to retrieve data for.'),
			opt('end_date', 'Latest date to retrieve data for; start_date if omitted.'),
			opt('interval', 'Data interval: 0 or omitted daily, 1 weekly, 2 monthly.'),
			opt(
				'headers',
				'Headers: 0 none, 1 or omitted show headers, 2 instrument identifier and headers.',
			),
			opt('property0', 'Column to return: 0 date, 1 close, 2 open, 3 high, 4 low, 5 volume.'),
			more('property1', 'More columns to return, in order.'),
		],
	),
	fn('SYD', "Calculates depreciation for one period with the sum-of-years'-digits method.", [
		COST,
		SALVAGE,
		LIFE,
		arg('per', 'The period to calculate, in the same units as life.'),
	]),
	fn(
		'VDB',
		'Calculates declining balance depreciation over any span of periods, including partial ones.',
		[
			COST,
			SALVAGE,
			LIFE,
			arg('start_period', 'Start of the period to calculate, in the same units as life.'),
			arg('end_period', 'End of the period to calculate.'),
			FACTOR,
			opt(
				'no_switch',
				'FALSE or omitted switches to straight-line when it is larger; TRUE never switches.',
			),
		],
	),
	fn(
		'XIRR',
		'Finds the annual rate at which dated, irregular cash flows have a net present value of zero.',
		[
			arg('values', 'Cash flows matching the dates; the first is usually the initial investment.'),
			arg('dates', 'Payment dates; the first one starts the schedule.'),
			GUESS,
		],
	),
	fn(
		'XNPV',
		'Discounts dated, irregular cash flows to today and adds them up (net present value).',
		[
			arg('rate', 'Discount rate to apply.'),
			arg('values', 'Cash flows matching the payment dates.'),
			arg('dates', 'Payment dates matching the cash flows.'),
		],
	),
];
