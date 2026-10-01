/** Cube functions: members, sets and values of an OLAP cube behind a workbook connection. */
import { arg, categoryBuilder, more, opt } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Cube');

const CONNECTION = arg('connection', 'Name of the workbook connection to the cube, as text.');
const CAPTION = opt('caption', 'Text to show in the cell instead of the caption from the cube.');

export const CUBE_INFO: readonly FunctionInfo[] = [
	fn(
		'CUBEKPIMEMBER',
		'Fetches a property of a key performance indicator (KPI) from a cube, showing the KPI name.',
		[
			CONNECTION,
			arg('kpi_name', 'Name of the KPI in the cube.'),
			arg(
				'kpi_property',
				'KPI part: 1 value, 2 goal, 3 status, 4 trend, 5 weight, 6 current time member.',
			),
			CAPTION,
		],
	),
	fn(
		'CUBEMEMBER',
		'Looks up a member or tuple in a cube; the result is an error when it does not exist.',
		[
			CONNECTION,
			arg('member_expression', 'An MDX expression that evaluates to a unique member or tuple.'),
			CAPTION,
		],
	),
	fn('CUBEMEMBERPROPERTY', 'Fetches the value of a named property of a member in a cube.', [
		CONNECTION,
		arg('member_expression', 'An MDX expression for a member of the cube.'),
		arg('property', 'Name of the property to return.'),
	]),
	fn('CUBERANKEDMEMBER', 'Picks the member at a given position of a set, such as the top seller.', [
		CONNECTION,
		arg('set_expression', 'A set expression, such as "{[Item1].children}", or a CUBESET cell.'),
		arg('rank', 'Position of the member to return; 1 is the first.'),
		CAPTION,
	]),
	fn('CUBESET', 'Asks the cube server to build a set of members or tuples from a set expression.', [
		CONNECTION,
		arg('set_expression', 'An MDX expression that evaluates to a set of members or tuples.'),
		CAPTION,
		opt(
			'sort_order',
			'Sort: 0 or omitted none, 1 ascending, 2 descending, 3 to 6 alphabetic or natural.',
		),
		opt('sort_by', 'The value to sort by, such as a measure.'),
	]),
	fn('CUBESETCOUNT', 'Counts the items of a set created with CUBESET.', [
		arg('set', 'A CUBESET expression, or a cell holding one.'),
	]),
	fn('CUBEVALUE', 'Fetches an aggregated value from a cube, sliced by member expressions.', [
		CONNECTION,
		opt('member_expression1', 'An MDX expression for a member or tuple that slices the cube.'),
		more('member_expression2', 'More member expressions to slice by.'),
	]),
];
