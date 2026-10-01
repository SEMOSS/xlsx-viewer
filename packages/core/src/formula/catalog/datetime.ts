/** Date & Time functions. */
import { BASIS } from './common-args';
import { arg, categoryBuilder, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Date & Time');

const DATE_SERIAL = arg('serial_number', 'The date to examine, as a serial number or date text.');
const TIME_SERIAL = arg('serial_number', 'The time to examine, as a serial number or time text.');
const START_DATE = arg('start_date', 'The starting date.');
const END_DATE = arg('end_date', 'The ending date.');
const HOLIDAYS = opt('holidays', 'Dates to leave out, such as public holidays.');
const WEEKEND = opt(
	'weekend',
	'Weekend days: a code such as 1 or omitted (Saturday and Sunday) or a mask such as "0000011".',
);
const WORKDAYS = arg('days', 'Working days to move; negative values go back in time.');
const MONTHS = arg('months', 'Months to move; negative values go back in time.');

export const DATETIME_INFO: readonly FunctionInfo[] = [
	fn('DATE', 'Returns the serial number of the date built from a year, month and day.', [
		arg('year', 'The year, from 1900 to 9999; smaller values are added to 1900.'),
		arg('month', 'Month of the year; values outside 1 to 12 roll into other years.'),
		arg('day', 'Day of the month; values outside the month roll into other months.'),
	]),
	fn('DATEDIF', 'Returns the number of whole days, months or years between two dates.', [
		arg('start_date', 'The earlier date.'),
		arg('end_date', 'The later date.'),
		arg('unit', 'What to count: "Y", "M", "D", "MD", "YM" or "YD".'),
	]),
	fn('DATEVALUE', 'Converts a date written as text into a date serial number.', [
		arg('date_text', 'Text that represents a date, such as "2024-03-15".'),
	]),
	fn('DAY', 'Extracts the day of the month, from 1 to 31, from a date.', [DATE_SERIAL]),
	fn('DAYS', 'Counts the days from one date to another.', [
		arg('end_date', 'The date to count up to.'),
		arg('start_date', 'The date to count from.'),
	]),
	fn(
		'DAYS360',
		'Counts the days between two dates as if every month had 30 days (a 360-day year).',
		[
			START_DATE,
			END_DATE,
			opt('method', 'FALSE or omitted uses the US (NASD) method; TRUE uses the European method.'),
		],
	),
	fn(
		'EDATE',
		'Moves a date forward or back by whole months and returns the resulting date serial.',
		[START_DATE, MONTHS],
	),
	fn(
		'EOMONTH',
		'Returns the last day of the month that lies a given number of months from a date.',
		[START_DATE, MONTHS],
	),
	fn('HOUR', 'Extracts the hour, from 0 to 23, from a time value.', [TIME_SERIAL]),
	fn('ISOWEEKNUM', 'Gives the ISO 8601 week of the year (1 to 53) that a date falls in.', [
		arg('date', 'The date to examine.'),
	]),
	fn('MINUTE', 'Extracts the minute, from 0 to 59, from a time value.', [TIME_SERIAL]),
	fn('MONTH', 'Extracts the month number, from 1 to 12, from a date.', [DATE_SERIAL]),
	fn(
		'NETWORKDAYS',
		'Counts the working days between two dates, leaving out weekends and holidays.',
		[START_DATE, END_DATE, HOLIDAYS],
	),
	fn('NETWORKDAYS.INTL', 'Counts the working days between two dates, with custom weekend days.', [
		START_DATE,
		END_DATE,
		WEEKEND,
		HOLIDAYS,
	]),
	fn('NOW', 'Gives the current date and time as a serial number that updates on recalculation.'),
	fn('SECOND', 'Extracts the second, from 0 to 59, from a time value.', [TIME_SERIAL]),
	fn('TIME', 'Returns the decimal number of a time built from hours, minutes and seconds.', [
		arg('hour', 'Hours, from 0 to 32767; whole days are dropped.'),
		arg('minute', 'Minutes; values over 59 carry into hours.'),
		arg('second', 'Seconds; values over 59 carry into minutes.'),
	]),
	fn('TIMEVALUE', 'Converts a time written as text into a decimal number from 0 to 1.', [
		arg('time_text', 'Text that represents a time, such as "6:45 PM".'),
	]),
	fn('TODAY', 'Gives the current date as a serial number that updates on recalculation.'),
	fn('WEEKDAY', 'Returns the day of the week of a date as a number.', [
		DATE_SERIAL,
		opt(
			'return_type',
			'Numbering: 1 or omitted Sunday=1, 2 Monday=1, 3 Monday=0; 11 to 17 also exist.',
		),
	]),
	fn('WEEKNUM', 'Returns the week number of a date within its year.', [
		DATE_SERIAL,
		opt('return_type', 'First day of the week: 1 or omitted Sunday, 2 Monday; 21 gives ISO weeks.'),
	]),
	fn('WORKDAY', 'Moves a date forward or back by working days, skipping weekends and holidays.', [
		START_DATE,
		WORKDAYS,
		HOLIDAYS,
	]),
	fn('WORKDAY.INTL', 'Moves a date forward or back by working days, with custom weekend days.', [
		START_DATE,
		WORKDAYS,
		WEEKEND,
		HOLIDAYS,
	]),
	fn('YEAR', 'Extracts the year, from 1900 to 9999, from a date.', [DATE_SERIAL]),
	fn('YEARFRAC', 'Returns the fraction of a year between two dates, using a day count basis.', [
		START_DATE,
		END_DATE,
		BASIS,
	]),
];
