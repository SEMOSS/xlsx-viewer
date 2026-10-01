/**
 * Financial functions, part 2: bonds, coupons, discounted securities and
 * Treasury bills.
 */
import { BASIS } from './common-args';
import { arg, categoryBuilder, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Financial');

const SETTLEMENT = arg(
	'settlement',
	'Settlement date: when the buyer takes ownership of the security.',
);
const MATURITY = arg('maturity', 'Maturity date: when the security expires.');
const ISSUE = arg('issue', 'Issue date of the security.');
const FREQUENCY = arg(
	'frequency',
	'Coupon payments per year: 1 annual, 2 semiannual, 4 quarterly.',
);
const COUPON = arg('rate', 'Annual coupon rate of the security.');
const RATE_AT_ISSUE = arg('rate', 'Interest rate of the security at the issue date.');
const YLD = arg('yld', 'Annual yield of the security.');
const PR = arg('pr', 'Price of the security per 100 of face value.');
const REDEMPTION = arg('redemption', 'Redemption value per 100 of face value.');
const DISCOUNT = arg('discount', 'Discount rate of the security.');
const INVESTMENT = arg('investment', 'Amount invested in the security.');
const FIRST_COUPON = arg('first_coupon', 'First coupon date of the security.');
const LAST_INTEREST = arg('last_interest', 'Last coupon date of the security.');

/** `settlement, maturity, frequency, [basis]` of the COUP* functions. */
const COUPON_ARGS = [SETTLEMENT, MATURITY, FREQUENCY, BASIS];

/** `settlement, maturity, coupon, yld, frequency, [basis]` of the duration functions. */
const DURATION_ARGS = [
	SETTLEMENT,
	MATURITY,
	arg('coupon', 'Annual coupon rate of the security.'),
	YLD,
	FREQUENCY,
	BASIS,
];

const TBILL_SETTLEMENT = arg('settlement', 'Settlement date of the Treasury bill.');
const TBILL_MATURITY = arg(
	'maturity',
	'Maturity date of the Treasury bill, within a year of settlement.',
);
const TBILL_DISCOUNT = arg('discount', 'Discount rate of the Treasury bill.');

export const SECURITIES_INFO: readonly FunctionInfo[] = [
	fn(
		'ACCRINT',
		'Calculates the interest a periodic-coupon security has accrued up to settlement.',
		[
			ISSUE,
			arg('first_interest', 'Date of the first interest payment.'),
			SETTLEMENT,
			COUPON,
			arg('par', 'Par value of the security.'),
			FREQUENCY,
			BASIS,
			opt(
				'calc_method',
				'TRUE or omitted accrues from the issue date; FALSE from the first interest date.',
			),
		],
	),
	fn('ACCRINTM', 'Calculates the interest accrued by a security that pays all of it at maturity.', [
		ISSUE,
		arg('settlement', 'Maturity date of the security.'),
		COUPON,
		arg('par', 'Par value of the security.'),
		BASIS,
	]),
	fn(
		'COUPDAYBS',
		'Counts the days from the start of the current coupon period up to settlement.',
		COUPON_ARGS,
	),
	fn('COUPDAYS', 'Counts the days in the coupon period that settlement falls in.', COUPON_ARGS),
	fn('COUPDAYSNC', 'Counts the days from settlement to the next coupon payment.', COUPON_ARGS),
	fn('COUPNCD', 'Gives the date of the next coupon payment after settlement.', COUPON_ARGS),
	fn(
		'COUPNUM',
		'Counts the coupon payments still due between settlement and maturity.',
		COUPON_ARGS,
	),
	fn('COUPPCD', 'Gives the date of the last coupon payment on or before settlement.', COUPON_ARGS),
	fn(
		'DISC',
		"Calculates the annual discount rate implied by a security's price and redemption value.",
		[SETTLEMENT, MATURITY, PR, REDEMPTION, BASIS],
	),
	fn(
		'DURATION',
		'Calculates the Macaulay duration, in years, of a security that pays periodic coupons.',
		DURATION_ARGS,
	),
	fn(
		'INTRATE',
		'Calculates the annual interest rate earned by a fully invested security held to maturity.',
		[
			SETTLEMENT,
			MATURITY,
			INVESTMENT,
			arg('redemption', 'Amount to be received at maturity.'),
			BASIS,
		],
	),
	fn(
		'MDURATION',
		'Calculates the modified duration (Macaulay duration adjusted for yield) of a security.',
		DURATION_ARGS,
	),
	fn(
		'ODDFPRICE',
		'Calculates the price of a security whose first coupon period is shorter or longer than usual.',
		[
			SETTLEMENT,
			MATURITY,
			ISSUE,
			FIRST_COUPON,
			arg('rate', 'Interest rate of the security.'),
			YLD,
			REDEMPTION,
			FREQUENCY,
			BASIS,
		],
	),
	fn(
		'ODDFYIELD',
		'Calculates the yield of a security whose first coupon period is shorter or longer than usual.',
		[
			SETTLEMENT,
			MATURITY,
			ISSUE,
			FIRST_COUPON,
			arg('rate', 'Interest rate of the security.'),
			PR,
			REDEMPTION,
			FREQUENCY,
			BASIS,
		],
	),
	fn(
		'ODDLPRICE',
		'Calculates the price of a security whose last coupon period is shorter or longer than usual.',
		[
			SETTLEMENT,
			MATURITY,
			LAST_INTEREST,
			arg('rate', 'Interest rate of the security.'),
			YLD,
			REDEMPTION,
			FREQUENCY,
			BASIS,
		],
	),
	fn(
		'ODDLYIELD',
		'Calculates the yield of a security whose last coupon period is shorter or longer than usual.',
		[
			SETTLEMENT,
			MATURITY,
			LAST_INTEREST,
			arg('rate', 'Interest rate of the security.'),
			PR,
			REDEMPTION,
			FREQUENCY,
			BASIS,
		],
	),
	fn(
		'PRICE',
		'Calculates the price of a periodic-coupon security from its yield (per 100 face value).',
		[SETTLEMENT, MATURITY, COUPON, YLD, REDEMPTION, FREQUENCY, BASIS],
	),
	fn(
		'PRICEDISC',
		'Calculates the price of a discounted (zero-coupon) security from its discount rate.',
		[SETTLEMENT, MATURITY, DISCOUNT, REDEMPTION, BASIS],
	),
	fn(
		'PRICEMAT',
		'Calculates the price of a security that pays all its interest at maturity, from its yield.',
		[SETTLEMENT, MATURITY, ISSUE, RATE_AT_ISSUE, YLD, BASIS],
	),
	fn(
		'RECEIVED',
		'Calculates the amount paid out at maturity for a fully invested discounted security.',
		[SETTLEMENT, MATURITY, INVESTMENT, DISCOUNT, BASIS],
	),
	fn('TBILLEQ', "Converts a Treasury bill's discount rate into its bond-equivalent yield.", [
		TBILL_SETTLEMENT,
		TBILL_MATURITY,
		TBILL_DISCOUNT,
	]),
	fn(
		'TBILLPRICE',
		'Calculates the price of a Treasury bill from its discount rate (per 100 face value).',
		[TBILL_SETTLEMENT, TBILL_MATURITY, TBILL_DISCOUNT],
	),
	fn('TBILLYIELD', 'Calculates the yield of a Treasury bill from its price.', [
		TBILL_SETTLEMENT,
		TBILL_MATURITY,
		arg('pr', 'Price of the Treasury bill per 100 of face value.'),
	]),
	fn('YIELD', 'Calculates the yield of a periodic-coupon security from its price.', [
		SETTLEMENT,
		MATURITY,
		COUPON,
		PR,
		REDEMPTION,
		FREQUENCY,
		BASIS,
	]),
	fn(
		'YIELDDISC',
		'Calculates the annual yield of a zero-coupon (discounted) security from its price.',
		[SETTLEMENT, MATURITY, PR, REDEMPTION, BASIS],
	),
	fn(
		'YIELDMAT',
		'Calculates the annual yield of a security that pays its interest at maturity, from its price.',
		[SETTLEMENT, MATURITY, ISSUE, RATE_AT_ISSUE, PR, BASIS],
	),
];
