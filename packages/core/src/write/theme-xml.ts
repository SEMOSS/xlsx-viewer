/**
 * `xl/theme/theme1.xml` generated from the Theme model: color scheme and
 * font scheme from the model, the standard Office (2013+) format scheme.
 */
import type { Theme, ThemeColorScheme } from '../types/theme';
import { XML_DECLARATION } from '../xml';
import { NS } from '../xml/namespaces';
import { attrString } from './xml-utils';

const SLOTS: Array<keyof Omit<ThemeColorScheme, 'name'>> = [
	'dk1',
	'lt1',
	'dk2',
	'lt2',
	'accent1',
	'accent2',
	'accent3',
	'accent4',
	'accent5',
	'accent6',
	'hlink',
	'folHlink',
];

function hex(value: string | undefined, fallback: string): string {
	const v = (value ?? '').replace(/^#/, '').toUpperCase();
	if (/^[0-9A-F]{8}$/.test(v)) {
		return v.slice(2);
	}
	return /^[0-9A-F]{6}$/.test(v) ? v : fallback;
}

function slotXml(slot: (typeof SLOTS)[number], value: string): string {
	if (slot === 'dk1' && value === '000000') {
		return '<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>';
	}
	if (slot === 'lt1' && value === 'FFFFFF') {
		return '<a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>';
	}
	return `<a:${slot}><a:srgbClr val="${value}"/></a:${slot}>`;
}

const DEFAULTS: Record<(typeof SLOTS)[number], string> = {
	dk1: '000000',
	lt1: 'FFFFFF',
	dk2: '44546A',
	lt2: 'E7E6E6',
	accent1: '4472C4',
	accent2: 'ED7D31',
	accent3: 'A5A5A5',
	accent4: 'FFC000',
	accent5: '5B9BD5',
	accent6: '70AD47',
	hlink: '0563C1',
	folHlink: '954F72',
};

/** Fallback faces for other scripts, as in the Office theme. */
const MAJOR_SCRIPTS: Array<[string, string]> = [
	['Jpan', '游ゴシック Light'],
	['Hang', '맑은 고딕'],
	['Hans', '等线 Light'],
	['Hant', '新細明體'],
	['Arab', 'Times New Roman'],
	['Hebr', 'Times New Roman'],
	['Thai', 'Tahoma'],
	['Ethi', 'Nyala'],
	['Beng', 'Vrinda'],
	['Gujr', 'Shruti'],
	['Khmr', 'MoolBoran'],
	['Knda', 'Tunga'],
	['Guru', 'Raavi'],
	['Cans', 'Euphemia'],
	['Cher', 'Plantagenet Cherokee'],
	['Yiii', 'Microsoft Yi Baiti'],
	['Tibt', 'Microsoft Himalaya'],
	['Thaa', 'MV Boli'],
	['Deva', 'Mangal'],
	['Telu', 'Gautami'],
	['Taml', 'Latha'],
	['Syrc', 'Estrangelo Edessa'],
	['Orya', 'Kalinga'],
	['Mlym', 'Kartika'],
	['Laoo', 'DokChampa'],
	['Sinh', 'Iskoola Pota'],
	['Mong', 'Mongolian Baiti'],
	['Viet', 'Times New Roman'],
	['Uigh', 'Microsoft Uighur'],
	['Geor', 'Sylfaen'],
];

const MINOR_OVERRIDES: Record<string, string> = {
	Jpan: '游ゴシック',
	Hans: '等线',
	Arab: 'Arial',
	Hebr: 'Arial',
	Khmr: 'DaunPenh',
	Viet: 'Arial',
};

function fontXml(
	tag: string,
	latin: string,
	ea: string | undefined,
	cs: string | undefined,
	minor: boolean,
): string {
	const scripts = MAJOR_SCRIPTS.map(([script, face]) => {
		const typeface = minor ? (MINOR_OVERRIDES[script] ?? face) : face;
		return `<a:font${attrString({ script, typeface })}/>`;
	}).join('');
	return (
		`<a:${tag}><a:latin${attrString({ typeface: latin })}/>` +
		`<a:ea${attrString({ typeface: ea ?? '' })}/><a:cs${attrString({ typeface: cs ?? '' })}/>` +
		`${scripts}</a:${tag}>`
	);
}

/** The Office format scheme (fills, lines, effects, background fills). */
const FMT_SCHEME =
	'<a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
	'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/>' +
	'<a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr">' +
	'<a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000">' +
	'<a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr>' +
	'</a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst>' +
	'<a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/>' +
	'</a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/>' +
	'<a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/>' +
	'<a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/>' +
	'</a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill>' +
	'<a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
	'<a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
	'<a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr">' +
	'<a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>' +
	'</a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/>' +
	'</a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" ' +
	'rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst>' +
	'</a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill>' +
	'<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill>' +
	'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/>' +
	'<a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000">' +
	'<a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/>' +
	'<a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/>' +
	'<a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>' +
	'</a:bgFillStyleLst></a:fmtScheme>';

/** Serialize a complete theme part. */
export function buildThemeXml(theme: Theme): string {
	const colors = SLOTS.map((s) => slotXml(s, hex(theme.colors[s], DEFAULTS[s]))).join('');
	const f = theme.fonts;
	return (
		XML_DECLARATION +
		`<a:theme xmlns:a="${NS.a}"${attrString({ name: theme.name || 'Office Theme' })}><a:themeElements>` +
		`<a:clrScheme${attrString({ name: theme.colors.name || 'Office' })}>${colors}</a:clrScheme>` +
		`<a:fontScheme${attrString({ name: f.name || 'Office' })}>` +
		fontXml(
			'majorFont',
			f.majorLatin || 'Calibri Light',
			f.majorEastAsian,
			f.majorComplexScript,
			false,
		) +
		fontXml('minorFont', f.minorLatin || 'Calibri', f.minorEastAsian, f.minorComplexScript, true) +
		`</a:fontScheme>${FMT_SCHEME}</a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>`
	);
}
