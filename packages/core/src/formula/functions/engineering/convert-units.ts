/**
 * Unit table for CONVERT (Excel 2013+ units). Each unit is expressed in a
 * base unit of its category: grams, meters, seconds, pascals, newtons,
 * joules, watts, teslas, cubic meters, square meters, bits and meters per
 * second. Temperatures are handled separately (they have offsets).
 */

export type Category =
	| 'mass'
	| 'distance'
	| 'time'
	| 'pressure'
	| 'force'
	| 'energy'
	| 'power'
	| 'magnetism'
	| 'temperature'
	| 'volume'
	| 'area'
	| 'information'
	| 'speed';

export interface UnitDef {
	category: Category;
	/** Size of one unit in the category's base unit (temperature: 1 = Kelvin step, see below). */
	factor: number;
	/** Metric prefixes allowed (`km`, `mg`); information units also take binary prefixes (`Kibyte`). */
	prefix?: boolean;
	/** Dimension the prefix applies to: 2 for areas (`km2` = 10^6 m2), 3 for volumes. */
	power?: 1 | 2 | 3;
}

const IN = 0.0254;
const FT = 0.3048;
const MI = 1609.344;
const NMI = 1852;
const LY = 9_460_730_472_580_800;
const PICA_PT = IN / 72;
const SURVEY_FT = 1200 / 3937;
const GAL = 3.785411784e-3;
const UK_GAL = 4.54609e-3;
const LBF = 4.4482216152605;
const HP = 550 * FT * LBF;

const table: Array<[string[], UnitDef]> = [
	// Weight and mass (grams)
	[['g'], { category: 'mass', factor: 1, prefix: true }],
	[['sg'], { category: 'mass', factor: (LBF / FT) * 1000 }],
	[['lbm'], { category: 'mass', factor: 453.59237 }],
	[['u'], { category: 'mass', factor: 1.6605390666e-24, prefix: true }],
	[['ozm'], { category: 'mass', factor: 453.59237 / 16 }],
	[['grain'], { category: 'mass', factor: 0.06479891 }],
	[['cwt', 'shweight'], { category: 'mass', factor: 45359.237 }],
	[['uk_cwt', 'lcwt', 'hweight'], { category: 'mass', factor: 50802.34544 }],
	[['stone'], { category: 'mass', factor: 6350.29318 }],
	[['ton'], { category: 'mass', factor: 907184.74 }],
	[['uk_ton', 'LTON', 'brton'], { category: 'mass', factor: 1016046.9088 }],
	// Distance (meters)
	[['m'], { category: 'distance', factor: 1, prefix: true }],
	[['mi'], { category: 'distance', factor: MI }],
	[['Nmi'], { category: 'distance', factor: NMI }],
	[['in'], { category: 'distance', factor: IN }],
	[['ft'], { category: 'distance', factor: FT }],
	[['yd'], { category: 'distance', factor: 0.9144 }],
	[['ang'], { category: 'distance', factor: 1e-10, prefix: true }],
	[['ell'], { category: 'distance', factor: 1.143 }],
	[['ly'], { category: 'distance', factor: LY, prefix: true }],
	[['parsec', 'pc'], { category: 'distance', factor: 3.0856775814913673e16, prefix: true }],
	[['Picapt', 'Pica'], { category: 'distance', factor: PICA_PT }],
	[['pica'], { category: 'distance', factor: IN / 6 }],
	[['survey_mi'], { category: 'distance', factor: 5280 * SURVEY_FT }],
	// Time (seconds)
	[['yr'], { category: 'time', factor: 365.25 * 86400 }],
	[['day', 'd'], { category: 'time', factor: 86400 }],
	[['hr'], { category: 'time', factor: 3600 }],
	[['mn', 'min'], { category: 'time', factor: 60 }],
	[['sec', 's'], { category: 'time', factor: 1, prefix: true }],
	// Pressure (pascals)
	[['Pa', 'p'], { category: 'pressure', factor: 1, prefix: true }],
	[['atm', 'at'], { category: 'pressure', factor: 101325, prefix: true }],
	[['mmHg'], { category: 'pressure', factor: 133.322, prefix: true }],
	[['psi'], { category: 'pressure', factor: LBF / (IN * IN) }],
	[['Torr'], { category: 'pressure', factor: 101325 / 760 }],
	// Force (newtons)
	[['N'], { category: 'force', factor: 1, prefix: true }],
	[['dyn', 'dy'], { category: 'force', factor: 1e-5, prefix: true }],
	[['lbf'], { category: 'force', factor: LBF }],
	[['pond'], { category: 'force', factor: 9.80665e-3, prefix: true }],
	// Energy (joules)
	[['J'], { category: 'energy', factor: 1, prefix: true }],
	[['e'], { category: 'energy', factor: 1e-7, prefix: true }],
	[['c'], { category: 'energy', factor: 4.184, prefix: true }],
	[['cal'], { category: 'energy', factor: 4.1868, prefix: true }],
	[['eV', 'ev'], { category: 'energy', factor: 1.602176634e-19, prefix: true }],
	[['HPh', 'hh'], { category: 'energy', factor: HP * 3600 }],
	[['Wh', 'wh'], { category: 'energy', factor: 3600, prefix: true }],
	// Excel's "foot-pound" is the foot-poundal (1 lb * 1 ft^2 / s^2).
	[['flb'], { category: 'energy', factor: 0.45359237 * FT * FT }],
	[['BTU', 'btu'], { category: 'energy', factor: 1055.05585262 }],
	// Power (watts)
	[['HP', 'h'], { category: 'power', factor: HP }],
	[['PS'], { category: 'power', factor: 735.49875 }],
	[['W', 'w'], { category: 'power', factor: 1, prefix: true }],
	// Magnetism (teslas)
	[['T'], { category: 'magnetism', factor: 1, prefix: true }],
	[['ga'], { category: 'magnetism', factor: 1e-4, prefix: true }],
	// Volume (cubic meters)
	[['tsp'], { category: 'volume', factor: GAL / 768 }],
	[['tspm'], { category: 'volume', factor: 5e-6 }],
	[['tbs'], { category: 'volume', factor: GAL / 256 }],
	[['oz'], { category: 'volume', factor: GAL / 128 }],
	[['cup'], { category: 'volume', factor: GAL / 16 }],
	[['pt', 'us_pt'], { category: 'volume', factor: GAL / 8 }],
	[['uk_pt'], { category: 'volume', factor: UK_GAL / 8 }],
	[['qt'], { category: 'volume', factor: GAL / 4 }],
	[['uk_qt'], { category: 'volume', factor: UK_GAL / 4 }],
	[['gal'], { category: 'volume', factor: GAL }],
	[['uk_gal'], { category: 'volume', factor: UK_GAL }],
	[['l', 'L', 'lt'], { category: 'volume', factor: 1e-3, prefix: true }],
	[['ang3', 'ang^3'], { category: 'volume', factor: 1e-30, prefix: true, power: 3 }],
	[['barrel'], { category: 'volume', factor: 42 * GAL }],
	[['bushel'], { category: 'volume', factor: 0.03523907016688 }],
	[['ft3', 'ft^3'], { category: 'volume', factor: FT ** 3 }],
	[['in3', 'in^3'], { category: 'volume', factor: IN ** 3 }],
	[['ly3', 'ly^3'], { category: 'volume', factor: LY ** 3, prefix: true, power: 3 }],
	[['m3', 'm^3'], { category: 'volume', factor: 1, prefix: true, power: 3 }],
	[['mi3', 'mi^3'], { category: 'volume', factor: MI ** 3 }],
	[['yd3', 'yd^3'], { category: 'volume', factor: 0.9144 ** 3 }],
	[['Nmi3', 'Nmi^3'], { category: 'volume', factor: NMI ** 3 }],
	[['Picapt3', 'Picapt^3', 'Pica3', 'Pica^3'], { category: 'volume', factor: PICA_PT ** 3 }],
	[['GRT', 'regton'], { category: 'volume', factor: 100 * FT ** 3 }],
	[['MTON'], { category: 'volume', factor: 40 * FT ** 3 }],
	// Area (square meters)
	[['uk_acre'], { category: 'area', factor: 4046.8564224 }],
	[['us_acre'], { category: 'area', factor: 43560 * SURVEY_FT * SURVEY_FT }],
	[['ang2', 'ang^2'], { category: 'area', factor: 1e-20, prefix: true, power: 2 }],
	[['ar'], { category: 'area', factor: 100, prefix: true }],
	[['ft2', 'ft^2'], { category: 'area', factor: FT * FT }],
	[['ha'], { category: 'area', factor: 10000 }],
	[['in2', 'in^2'], { category: 'area', factor: IN * IN }],
	[['ly2', 'ly^2'], { category: 'area', factor: LY * LY, prefix: true, power: 2 }],
	[['m2', 'm^2'], { category: 'area', factor: 1, prefix: true, power: 2 }],
	[['Morgen'], { category: 'area', factor: 2500 }],
	[['mi2', 'mi^2'], { category: 'area', factor: MI * MI }],
	[['Nmi2', 'Nmi^2'], { category: 'area', factor: NMI * NMI }],
	[['Picapt2', 'Pica2', 'Pica^2', 'Picapt^2'], { category: 'area', factor: PICA_PT * PICA_PT }],
	[['yd2', 'yd^2'], { category: 'area', factor: 0.9144 * 0.9144 }],
	// Information (bits)
	[['bit'], { category: 'information', factor: 1, prefix: true }],
	[['byte'], { category: 'information', factor: 8, prefix: true }],
	// Speed (meters per second)
	[['admkn'], { category: 'speed', factor: (6080 * FT) / 3600 }],
	[['kn'], { category: 'speed', factor: NMI / 3600 }],
	[['m/h', 'm/hr'], { category: 'speed', factor: 1 / 3600, prefix: true }],
	[['m/s', 'm/sec'], { category: 'speed', factor: 1, prefix: true }],
	[['mph'], { category: 'speed', factor: MI / 3600 }],
	// Temperature: factor is unused, see TEMPERATURES.
	[['C', 'cel'], { category: 'temperature', factor: 1 }],
	[['F', 'fah'], { category: 'temperature', factor: 1 }],
	[['K', 'kel'], { category: 'temperature', factor: 1, prefix: true }],
	[['Rank'], { category: 'temperature', factor: 1 }],
	[['Reau'], { category: 'temperature', factor: 1 }],
];

/** Units by case-sensitive name. */
export const UNITS: ReadonlyMap<string, UnitDef> = new Map(
	table.flatMap(([names, def]) => names.map((n): [string, UnitDef] => [n, def])),
);

/**
 * Temperature scales as `celsius = (x - offset) * num / den`, kept rational
 * so that whole-degree conversions such as 68 F -> 20 C are exact.
 */
export const TEMPERATURES: ReadonlyMap<string, { offset: number; num: number; den: number }> =
	new Map([
		['C', { offset: 0, num: 1, den: 1 }],
		['cel', { offset: 0, num: 1, den: 1 }],
		['F', { offset: 32, num: 5, den: 9 }],
		['fah', { offset: 32, num: 5, den: 9 }],
		['K', { offset: 273.15, num: 1, den: 1 }],
		['kel', { offset: 273.15, num: 1, den: 1 }],
		['Rank', { offset: 491.67, num: 5, den: 9 }],
		['Reau', { offset: 0, num: 5, den: 4 }],
	]);

/** Decimal prefixes as powers of ten (`da` and `e` are both deka). */
export const DECIMAL_PREFIXES: ReadonlyMap<string, number> = new Map([
	['Y', 24],
	['Z', 21],
	['E', 18],
	['P', 15],
	['T', 12],
	['G', 9],
	['M', 6],
	['k', 3],
	['h', 2],
	['da', 1],
	['e', 1],
	['d', -1],
	['c', -2],
	['m', -3],
	['u', -6],
	['n', -9],
	['p', -12],
	['f', -15],
	['a', -18],
	['z', -21],
	['y', -24],
]);

/** Binary prefixes as powers of two (information units only). */
export const BINARY_PREFIXES: ReadonlyMap<string, number> = new Map([
	['Yi', 80],
	['Zi', 70],
	['Ei', 60],
	['Pi', 50],
	['Ti', 40],
	['Gi', 30],
	['Mi', 20],
	['ki', 10],
]);
