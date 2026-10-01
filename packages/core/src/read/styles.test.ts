import { describe, expect, it } from 'vitest';

import { serializeElement } from '../xml/serializer';
import { parseStyleSheet } from './styles';
import { parseTheme } from './theme';
import { MAIN_NS } from './test-utils/package-builder';

const STYLES = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="${MAIN_NS}" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" mc:Ignorable="x14ac">
	<numFmts count="2"><numFmt numFmtId="164" formatCode="0.000"/><numFmt numFmtId="165" formatCode="&quot;$&quot;#,##0.00"/></numFmts>
	<fonts count="3" x14ac:knownFonts="1">
		<font><sz val="11"/><color theme="1"/><name val="Calibri"/><family val="2"/><scheme val="minor"/></font>
		<font><b/><i val="0"/><strike/><u val="double"/><vertAlign val="superscript"/><sz val="14"/><color rgb="FF0070C0"/><name val="Arial"/><charset val="1"/><outline/><shadow/><condense val="0"/><extend/></font>
		<font><u/><sz val="9"/><color indexed="10" tint="-0.25"/><name val="Tahoma"/></font>
	</fonts>
	<fills count="5">
		<fill><patternFill patternType="none"/></fill>
		<fill><patternFill patternType="gray125"/></fill>
		<fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/><bgColor indexed="64"/></patternFill></fill>
		<fill><gradientFill degree="90"><stop position="0"><color theme="0"/></stop><stop position="1"><color theme="4" tint="0.5"/></stop></gradientFill></fill>
		<fill><gradientFill type="path" left="0.5" right="0.5" top="0.5" bottom="0.5"><stop position="0"><color rgb="FFFFFFFF"/></stop><stop position="1"><color rgb="FF000000"/></stop></gradientFill></fill>
	</fills>
	<borders count="2">
		<border><left/><right/><top/><bottom/><diagonal/></border>
		<border diagonalUp="1" diagonalDown="1"><left style="thin"><color indexed="64"/></left><right style="medium"><color auto="1"/></right><top style="dashed"/><bottom style="double"><color theme="3"/></bottom><diagonal style="hair"><color rgb="FFFF0000"/></diagonal></border>
	</borders>
	<cellStyleXfs count="2"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/><xf numFmtId="0" fontId="1" fillId="0" borderId="0" applyFont="0"/></cellStyleXfs>
	<cellXfs count="3">
		<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
		<xf numFmtId="164" fontId="1" fillId="2" borderId="1" xfId="1" applyNumberFormat="1" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1" applyProtection="1" quotePrefix="1" pivotButton="1"><alignment horizontal="center" vertical="top" wrapText="1" shrinkToFit="0" indent="2" textRotation="45" readingOrder="2" justifyLastLine="1" relativeIndent="1"/><protection locked="0" hidden="1"/></xf>
		<xf numFmtId="14" fontId="2" fillId="3" borderId="0" xfId="0"><alignment horizontal="centerContinuous"/></xf>
	</cellXfs>
	<cellStyles count="2"><cellStyle name="Normal" xfId="0" builtinId="0"/><cellStyle name="Heading 1" xfId="1" builtinId="16" iLevel="0" hidden="1" customBuiltin="1"/></cellStyles>
	<dxfs count="2">
		<dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></dxf>
		<dxf><numFmt numFmtId="166" formatCode="0.0%"/><alignment horizontal="right"/><border><left style="thin"/><vertical style="thin"/><horizontal style="dotted"/></border><protection locked="1"/></dxf>
	</dxfs>
	<tableStyles count="1" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16">
		<tableStyle name="MyStyle" pivot="0" count="2"><tableStyleElement type="wholeTable" dxfId="0"/><tableStyleElement type="firstRowStripe" dxfId="1" size="2"/></tableStyle>
	</tableStyles>
	<colors><indexedColors><rgbColor rgb="FF000000"/><rgbColor rgb="ffffff"/></indexedColors><mruColors><color rgb="FF00FF00"/></mruColors></colors>
	<extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}"/></extLst>
</styleSheet>`;

describe('parseStyleSheet', () => {
	const styles = parseStyleSheet(STYLES);

	it('reads number formats and fonts', () => {
		expect(styles.numFmts).toEqual({ 164: '0.000', 165: '"$"#,##0.00' });
		expect(styles.fonts[0]).toEqual({
			size: 11,
			color: { theme: 1 },
			name: 'Calibri',
			family: 2,
			scheme: 'minor',
		});
		expect(styles.fonts[1]).toEqual({
			bold: true,
			italic: false,
			strike: true,
			underline: 'double',
			vertAlign: 'superscript',
			size: 14,
			color: { rgb: 'FF0070C0' },
			name: 'Arial',
			charset: 1,
			outline: true,
			shadow: true,
			condense: false,
			extend: true,
		});
		expect(styles.fonts[2]).toEqual({
			underline: 'single',
			size: 9,
			color: { indexed: 10, tint: -0.25 },
			name: 'Tahoma',
		});
	});

	it('reads pattern and gradient fills', () => {
		expect(styles.fills[0]).toEqual({ type: 'pattern', patternType: 'none' });
		expect(styles.fills[1]).toEqual({ type: 'pattern', patternType: 'gray125' });
		expect(styles.fills[2]).toEqual({
			type: 'pattern',
			patternType: 'solid',
			fgColor: { rgb: 'FFFFFF00' },
			bgColor: { indexed: 64 },
		});
		expect(styles.fills[3]).toEqual({
			type: 'gradient',
			gradientType: 'linear',
			degree: 90,
			stops: [
				{ position: 0, color: { theme: 0 } },
				{ position: 1, color: { theme: 4, tint: 0.5 } },
			],
		});
		expect(styles.fills[4]).toMatchObject({
			gradientType: 'path',
			left: 0.5,
			right: 0.5,
			top: 0.5,
			bottom: 0.5,
		});
	});

	it('reads borders with every edge and diagonal flags', () => {
		expect(styles.borders[0]).toEqual({});
		expect(styles.borders[1]).toEqual({
			left: { style: 'thin', color: { indexed: 64 } },
			right: { style: 'medium', color: { auto: true } },
			top: { style: 'dashed' },
			bottom: { style: 'double', color: { theme: 3 } },
			diagonal: { style: 'hair', color: { rgb: 'FFFF0000' } },
			diagonalUp: true,
			diagonalDown: true,
		});
	});

	it('reads cell formats with alignment, protection and apply flags', () => {
		expect(styles.cellStyleXfs).toEqual([
			{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 },
			{ numFmtId: 0, fontId: 1, fillId: 0, borderId: 0, applyFont: false },
		]);
		expect(styles.cellXfs[1]).toEqual({
			numFmtId: 164,
			fontId: 1,
			fillId: 2,
			borderId: 1,
			xfId: 1,
			alignment: {
				horizontal: 'center',
				vertical: 'top',
				wrapText: true,
				shrinkToFit: false,
				indent: 2,
				textRotation: 45,
				readingOrder: 2,
				justifyLastLine: true,
				relativeIndent: 1,
			},
			protection: { locked: false, hidden: true },
			applyNumberFormat: true,
			applyFont: true,
			applyFill: true,
			applyBorder: true,
			applyAlignment: true,
			applyProtection: true,
			quotePrefix: true,
			pivotButton: true,
		});
		expect(styles.cellXfs[2].alignment).toEqual({ horizontal: 'centerContinuous' });
		expect(styles.cellStyles[1]).toEqual({
			name: 'Heading 1',
			xfId: 1,
			builtinId: 16,
			iLevel: 0,
			hidden: true,
			customBuiltin: true,
		});
	});

	it('reads differential formats (a dxf pattern without type is solid)', () => {
		expect(styles.dxfs[0]).toEqual({
			font: { bold: true, color: { rgb: 'FF9C0006' } },
			fill: { type: 'pattern', patternType: 'solid', bgColor: { rgb: 'FFFFC7CE' } },
		});
		expect(styles.dxfs[1]).toEqual({
			numFmt: { id: 166, code: '0.0%' },
			alignment: { horizontal: 'right' },
			border: {
				left: { style: 'thin' },
				vertical: { style: 'thin' },
				horizontal: { style: 'dotted' },
			},
			protection: { locked: true },
		});
	});

	it('reads table styles and the indexed palette', () => {
		expect(styles.defaultTableStyle).toBe('TableStyleMedium9');
		expect(styles.defaultPivotStyle).toBe('PivotStyleLight16');
		expect(styles.customTableStyles).toEqual([
			{
				name: 'MyStyle',
				pivot: false,
				table: true,
				elements: [
					{ type: 'wholeTable', dxfId: 0 },
					{ type: 'firstRowStripe', dxfId: 1, size: 2 },
				],
			},
		]);
		expect(styles.indexedColors).toEqual(['FF000000', 'FFFFFFFF']);
	});

	it('keeps the original element of every entry for verbatim re-emission', () => {
		const raw = styles.raw!;
		expect(raw.root?.name).toBe('styleSheet');
		expect(raw.fonts).toHaveLength(3);
		expect(raw.fills).toHaveLength(5);
		expect(raw.borders).toHaveLength(2);
		expect(raw.cellStyleXfs).toHaveLength(2);
		expect(raw.cellXfs).toHaveLength(3);
		expect(raw.cellStyles).toHaveLength(2);
		expect(raw.dxfs).toHaveLength(2);
		expect(serializeElement(raw.fonts![2])).toBe(
			'<font><u/><sz val="9"/><color indexed="10" tint="-0.25"/><name val="Tahoma"/></font>',
		);
		expect(serializeElement(raw.cellXfs![2])).toBe(
			'<xf numFmtId="14" fontId="2" fillId="3" borderId="0" xfId="0"><alignment horizontal="centerContinuous"/></xf>',
		);
	});

	it('pads a styles part that lacks the mandatory first entries', () => {
		const s = parseStyleSheet(`<styleSheet xmlns="${MAIN_NS}"><numFmts count="0"/></styleSheet>`);
		expect(s.fonts).toHaveLength(1);
		expect(s.fills).toHaveLength(2);
		expect(s.borders).toHaveLength(1);
		expect(s.cellXfs).toHaveLength(1);
		expect(s.raw?.cellXfs).toEqual([]);
	});

	it('reads a styles part whose elements use a namespace prefix', () => {
		const s = parseStyleSheet(
			`<x:styleSheet xmlns:x="${MAIN_NS}"><x:fonts count="1"><x:font><x:b/><x:sz val="10"/></x:font></x:fonts>` +
				`<x:cellXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></x:cellXfs></x:styleSheet>`,
		);
		expect(s.fonts).toEqual([{ bold: true, size: 10 }]);
		expect(s.raw?.root?.attrs.xmlns).toBe(MAIN_NS);
		expect(s.raw?.fonts?.[0].name).toBe('font');
	});
});

describe('parseTheme', () => {
	it('reads colors (srgb and system), fonts and keeps the XML', () => {
		const xml =
			'<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Custom"><a:themeElements>' +
			'<a:clrScheme name="Mine"><a:dk1><a:sysClr val="windowText" lastClr="111111"/></a:dk1><a:lt1><a:sysClr val="window"/></a:lt1>' +
			'<a:dk2><a:srgbClr val="44546a"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="FF0000"/></a:accent1>' +
			'<a:hlink><a:srgbClr val="0000FF"/></a:hlink></a:clrScheme>' +
			'<a:fontScheme name="Fonts"><a:majorFont><a:latin typeface="Georgia"/><a:ea typeface="MS Mincho"/><a:cs typeface=""/></a:majorFont>' +
			'<a:minorFont><a:latin typeface="Verdana"/><a:ea typeface=""/><a:cs typeface="Arial"/></a:minorFont></a:fontScheme>' +
			'</a:themeElements></a:theme>';
		const theme = parseTheme(xml);
		expect(theme.name).toBe('Custom');
		expect(theme.colors).toMatchObject({
			name: 'Mine',
			dk1: '111111',
			lt1: 'FFFFFF',
			dk2: '44546A',
			accent1: 'FF0000',
			hlink: '0000FF',
			// Missing slots fall back to the default Office theme.
			accent2: 'ED7D31',
		});
		expect(theme.fonts).toEqual({
			name: 'Fonts',
			majorLatin: 'Georgia',
			minorLatin: 'Verdana',
			majorEastAsian: 'MS Mincho',
			minorComplexScript: 'Arial',
		});
		expect(theme.rawXml).toBe(xml);
	});
});
