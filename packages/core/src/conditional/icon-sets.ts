import type { IconSetName } from '../types/conditional-format';

/** Drawable icon shapes used by Excel's icon sets. */
export type IconShape =
	| 'arrowUp'
	| 'arrowDown'
	| 'arrowRight'
	| 'arrowUpRight'
	| 'arrowDownRight'
	| 'circle'
	| 'trafficLight'
	| 'flag'
	| 'triangleUp'
	| 'triangleDown'
	| 'diamond'
	| 'dash'
	| 'star'
	| 'starHalf'
	| 'starEmpty'
	| 'check'
	| 'cross'
	| 'exclamation'
	| 'quarter'
	| 'rating'
	| 'box';

/** How to draw one icon. */
export interface IconDescriptor {
	shape: IconShape;
	/** Main color (CSS hex). */
	color: string;
	/** Filled portion (0..1) for `quarter`, `rating` and `box` icons. */
	fill?: number;
	/** The symbol is drawn in white on a filled circle of `color` (3 Symbols, circled). */
	circled?: boolean;
}

export interface IconSetDescriptor {
	name: IconSetName;
	/** Name used by Excel's Icon Sets gallery. */
	label: string;
	/** Number of icons (thresholds). 0 for `NoIcons` ("No Cell Icon" in custom icon sets). */
	count: number;
	/** Icons from the lowest values (index 0) to the highest. */
	icons: readonly IconDescriptor[];
}

/** Icon palette (Excel 2013+ flat icons). */
export const ICON_COLORS = Object.freeze({
	red: '#E0393E',
	yellow: '#FFC000',
	green: '#00B050',
	gray: '#808080',
	black: '#404040',
	lightGray: '#A6A6A6',
	pink: '#F4A6A6',
	gold: '#FFC000',
	blue: '#2F75B5',
});

const C = ICON_COLORS;

function icon(
	shape: IconShape,
	color: string,
	extra: Partial<IconDescriptor> = {},
): IconDescriptor {
	return Object.freeze({ shape, color, ...extra });
}

function iconSet(name: IconSetName, label: string, icons: IconDescriptor[]): IconSetDescriptor {
	return Object.freeze({ name, label, count: icons.length, icons: Object.freeze(icons) });
}

function arrows(count: 3 | 4 | 5, gray: boolean): IconDescriptor[] {
	const red = gray ? C.gray : C.red;
	const yellow = gray ? C.gray : C.yellow;
	const green = gray ? C.gray : C.green;
	const down = icon('arrowDown', red);
	const up = icon('arrowUp', green);
	if (count === 3) {
		return [down, icon('arrowRight', yellow), up];
	}
	if (count === 4) {
		return [down, icon('arrowDownRight', yellow), icon('arrowUpRight', yellow), up];
	}
	return [
		down,
		icon('arrowDownRight', yellow),
		icon('arrowRight', yellow),
		icon('arrowUpRight', yellow),
		up,
	];
}

function filled(shape: IconShape, color: string, fills: number[]): IconDescriptor[] {
	return fills.map((fill) => icon(shape, color, { fill }));
}

/** Every Excel icon set, keyed by its OOXML name. */
export const ICON_SETS: Readonly<Record<IconSetName, IconSetDescriptor>> = Object.freeze({
	'3Arrows': iconSet('3Arrows', '3 Arrows (Colored)', arrows(3, false)),
	'3ArrowsGray': iconSet('3ArrowsGray', '3 Arrows (Gray)', arrows(3, true)),
	'3Flags': iconSet('3Flags', '3 Flags', [
		icon('flag', C.red),
		icon('flag', C.yellow),
		icon('flag', C.green),
	]),
	'3TrafficLights1': iconSet('3TrafficLights1', '3 Traffic Lights (Unrimmed)', [
		icon('circle', C.red),
		icon('circle', C.yellow),
		icon('circle', C.green),
	]),
	'3TrafficLights2': iconSet('3TrafficLights2', '3 Traffic Lights (Rimmed)', [
		icon('trafficLight', C.red),
		icon('trafficLight', C.yellow),
		icon('trafficLight', C.green),
	]),
	'3Signs': iconSet('3Signs', '3 Signs', [
		icon('diamond', C.red),
		icon('triangleUp', C.yellow),
		icon('circle', C.green),
	]),
	'3Symbols': iconSet('3Symbols', '3 Symbols (Circled)', [
		icon('cross', C.red, { circled: true }),
		icon('exclamation', C.yellow, { circled: true }),
		icon('check', C.green, { circled: true }),
	]),
	'3Symbols2': iconSet('3Symbols2', '3 Symbols (Uncircled)', [
		icon('cross', C.red),
		icon('exclamation', C.yellow),
		icon('check', C.green),
	]),
	'3Stars': iconSet('3Stars', '3 Stars', [
		icon('starEmpty', C.gold),
		icon('starHalf', C.gold),
		icon('star', C.gold),
	]),
	'3Triangles': iconSet('3Triangles', '3 Triangles', [
		icon('triangleDown', C.red),
		icon('dash', C.yellow),
		icon('triangleUp', C.green),
	]),
	'4Arrows': iconSet('4Arrows', '4 Arrows (Colored)', arrows(4, false)),
	'4ArrowsGray': iconSet('4ArrowsGray', '4 Arrows (Gray)', arrows(4, true)),
	'4RedToBlack': iconSet('4RedToBlack', 'Red To Black', [
		icon('circle', C.black),
		icon('circle', C.lightGray),
		icon('circle', C.pink),
		icon('circle', C.red),
	]),
	'4Rating': iconSet('4Rating', '4 Ratings', filled('rating', C.blue, [0.25, 0.5, 0.75, 1])),
	'4TrafficLights': iconSet('4TrafficLights', '4 Traffic Lights', [
		icon('circle', C.black),
		icon('circle', C.red),
		icon('circle', C.yellow),
		icon('circle', C.green),
	]),
	'5Arrows': iconSet('5Arrows', '5 Arrows (Colored)', arrows(5, false)),
	'5ArrowsGray': iconSet('5ArrowsGray', '5 Arrows (Gray)', arrows(5, true)),
	'5Rating': iconSet('5Rating', '5 Ratings', filled('rating', C.blue, [0, 0.25, 0.5, 0.75, 1])),
	'5Quarters': iconSet(
		'5Quarters',
		'5 Quarters',
		filled('quarter', C.black, [0, 0.25, 0.5, 0.75, 1]),
	),
	'5Boxes': iconSet('5Boxes', '5 Boxes', filled('box', C.blue, [0, 0.25, 0.5, 0.75, 1])),
	NoIcons: iconSet('NoIcons', 'No Cell Icon', []),
});

/** Icon sets grouped like Excel's Conditional Formatting > Icon Sets gallery. */
export const ICON_SET_GROUPS: ReadonlyArray<{ label: string; sets: readonly IconSetName[] }> =
	Object.freeze([
		{
			label: 'Directional',
			sets: [
				'3Arrows',
				'3ArrowsGray',
				'3Triangles',
				'4Arrows',
				'4ArrowsGray',
				'5Arrows',
				'5ArrowsGray',
			],
		},
		{
			label: 'Shapes',
			sets: ['3TrafficLights1', '3TrafficLights2', '3Signs', '4TrafficLights', '4RedToBlack'],
		},
		{ label: 'Indicators', sets: ['3Symbols', '3Symbols2', '3Flags'] },
		{ label: 'Ratings', sets: ['3Stars', '4Rating', '5Quarters', '5Rating', '5Boxes'] },
	]);
