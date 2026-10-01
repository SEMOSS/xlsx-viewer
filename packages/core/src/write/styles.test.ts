import { describe, expect, it } from 'vitest';

import { createDefaultStyleSheet } from '../styles/defaults';
import { StylePool } from '../styles/style-pool';
import { parseXml } from '../xml';
import { borderXml, dxfXml, fillXml, fontXml, xfXml } from './style-entries';
import { buildStylesXml, stylesChanged } from './styles-xml';
import { childNames, pick, pickAll } from './test-utils/package-check';

const SOURCE =
	'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">' +
	'<numFmts count="1"><numFmt numFmtId="164" formatCode="0.000"/></numFmts>' +
	'<fonts count="1" x14ac:knownFonts="1"><font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/><extra-unknown/></font></fonts>' +
	'<fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills>' +
	'<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>' +
	'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
	'<cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs>' +
	'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
	'<dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/>' +
	'<colors><mruColors><color rgb="FF00B050"/></mruColors></colors>' +
	'<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext></extLst>' +
	'</styleSheet>';

function loaded() {
	const root = parseXml(SOURCE).root;
	const styles = createDefaultStyleSheet();
	styles.numFmts = { 164: '0.000' };
	return { root, styles };
}

describe('style entries', () => {
	it('serializes fonts in Excel order', () => {
		expect(
			fontXml({
				bold: true,
				italic: false,
				underline: 'double',
				size: 12,
				color: { rgb: '00FF00' },
				name: 'Arial',
				family: 2,
				scheme: 'none',
			}),
		).toBe(
			'<font><b/><i val="0"/><u val="double"/><sz val="12"/><color rgb="FF00FF00"/><name val="Arial"/><family val="2"/><scheme val="none"/></font>',
		);
		expect(fontXml({})).toBe('<font/>');
	});

	it('serializes pattern and gradient fills', () => {
		expect(
			fillXml({ type: 'pattern', patternType: 'solid', fgColor: { theme: 4, tint: 0.4 } }),
		).toBe(
			'<fill><patternFill patternType="solid"><fgColor theme="4" tint="0.4"/></patternFill></fill>',
		);
		expect(
			fillXml({
				type: 'gradient',
				gradientType: 'linear',
				degree: 90,
				stops: [
					{ position: 0, color: { rgb: 'FFFFFFFF' } },
					{ position: 1, color: { theme: 4 } },
				],
			}),
		).toBe(
			'<fill><gradientFill degree="90"><stop position="0"><color rgb="FFFFFFFF"/></stop><stop position="1"><color theme="4"/></stop></gradientFill></fill>',
		);
	});

	it('writes solid dxf fills with both colors', () => {
		expect(
			dxfXml({
				fill: { type: 'pattern', patternType: 'solid', bgColor: { rgb: 'FFFFC7CE' } },
				font: { color: { rgb: 'FF9C0006' } },
			}),
		).toBe(
			'<dxf><font><color rgb="FF9C0006"/></font><fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>',
		);
	});

	it('serializes borders (all edges for cells, set edges for dxfs)', () => {
		const b = { top: { style: 'thin' as const, color: { indexed: 64 } }, diagonalUp: true };
		expect(borderXml(b)).toBe(
			'<border diagonalUp="1"><left/><right/><top style="thin"><color indexed="64"/></top><bottom/><diagonal/></border>',
		);
		expect(borderXml(b, true)).toBe(
			'<border diagonalUp="1"><top style="thin"><color indexed="64"/></top></border>',
		);
	});

	it('serializes xfs with alignment and protection', () => {
		expect(
			xfXml(
				{
					numFmtId: 164,
					fontId: 1,
					fillId: 2,
					borderId: 0,
					xfId: 0,
					applyNumberFormat: true,
					applyAlignment: true,
					alignment: { horizontal: 'center', wrapText: true, indent: 0 },
					protection: { locked: false },
				},
				true,
			),
		).toBe(
			'<xf numFmtId="164" fontId="1" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="center" wrapText="1"/><protection locked="0"/></xf>',
		);
	});
});

describe('styles.xml', () => {
	it('writes a complete stylesheet for a new workbook', () => {
		const root = parseXml(buildStylesXml(createDefaultStyleSheet(), undefined)).root;
		expect(childNames(root)).toEqual([
			'fonts',
			'fills',
			'borders',
			'cellStyleXfs',
			'cellXfs',
			'cellStyles',
			'dxfs',
			'tableStyles',
		]);
		expect(pick(root, 'fills')?.attrs.count).toBe('2');
		expect(pick(root, 'tableStyles')?.attrs.defaultTableStyle).toBe('TableStyleMedium2');
	});

	it('detects whether a loaded stylesheet grew', () => {
		const { root, styles } = loaded();
		expect(stylesChanged(styles, root)).toBe(false);
		new StylePool(styles).derive(0, { font: { bold: true } });
		expect(stylesChanged(styles, root)).toBe(true);
	});

	it('appends new entries after the original ones, keeping raw XML and unknown markup', () => {
		const { root, styles } = loaded();
		const pool = new StylePool(styles);
		const xf = pool.derive(0, {
			font: { bold: true },
			fill: { type: 'pattern', patternType: 'solid', fgColor: { rgb: 'FFFFFF00' } },
			numFmt: '0.0%',
		});
		styles.dxfs.push({ font: { italic: true } });
		const out = parseXml(buildStylesXml(styles, root)).root;
		expect(childNames(out)).toEqual([
			'numFmts',
			'fonts',
			'fills',
			'borders',
			'cellStyleXfs',
			'cellXfs',
			'cellStyles',
			'dxfs',
			'tableStyles',
			'colors',
			'extLst',
		]);
		const fonts = pick(out, 'fonts')!;
		expect(fonts.attrs).toMatchObject({ count: '2', 'x14ac:knownFonts': '1' });
		expect(childNames(pickAll(fonts, 'font')[0])).toContain('extra-unknown');
		expect(childNames(pickAll(fonts, 'font')[1])).toEqual([
			'b',
			'sz',
			'color',
			'name',
			'family',
			'scheme',
		]);
		expect(
			pickAll(pick(out, 'numFmts'), 'numFmt').map((n) => [n.attrs.numFmtId, n.attrs.formatCode]),
		).toEqual([
			['164', '0.000'],
			['165', '0.0%'],
		]);
		const xfs = pickAll(pick(out, 'cellXfs'), 'xf');
		expect(xfs).toHaveLength(2);
		expect(xf).toBe(1);
		expect(xfs[1].attrs).toMatchObject({
			numFmtId: '165',
			fontId: '1',
			fillId: '2',
			applyFont: '1',
			applyFill: '1',
		});
		expect(pick(out, 'dxfs')?.attrs.count).toBe('1');
		expect(pick(out, 'colors', 'mruColors', 'color')?.attrs.rgb).toBe('FF00B050');
		expect(pick(out, 'extLst', 'ext')?.attrs.uri).toBe('{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}');
	});

	it('writes custom table styles and indexed colors', () => {
		const { root, styles } = loaded();
		styles.customTableStyles = [
			{ name: 'Mine', pivot: false, elements: [{ type: 'wholeTable', dxfId: 0 }] },
		];
		styles.dxfs.push({ border: { top: { style: 'thin' } } });
		styles.indexedColors = ['000000', 'FFFFFF'];
		const out = parseXml(buildStylesXml(styles, root)).root;
		const ts = pick(out, 'tableStyles')!;
		expect(ts.attrs.count).toBe('1');
		expect(pick(ts, 'tableStyle')?.attrs).toEqual({ name: 'Mine', pivot: '0', count: '1' });
		expect(childNames(pick(out, 'colors'))).toEqual(['indexedColors', 'mruColors']);
		expect(
			pickAll(pick(out, 'colors', 'indexedColors'), 'rgbColor').map((c) => c.attrs.rgb),
		).toEqual(['FF000000', 'FFFFFFFF']);
	});
});
