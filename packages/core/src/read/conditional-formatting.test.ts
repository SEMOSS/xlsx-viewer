import { describe, expect, it } from 'vitest';

import type { Worksheet } from '../types/sheet';
import { readWorkbook } from './read-workbook';
import { X14_NS, XM_NS, buildWorkbook, worksheet } from './test-utils/package-builder';

const ID_BAR = '{A1B2C3D4-0000-0000-0000-000000000001}';
const ID_STARS = '{A1B2C3D4-0000-0000-0000-000000000002}';
const ID_EXPR = '{A1B2C3D4-0000-0000-0000-000000000003}';

const CF_2007 =
	'<conditionalFormatting sqref="A1:A10"><cfRule type="dataBar" priority="1"><dataBar showValue="0" minLength="5" maxLength="95">' +
	'<cfvo type="min"/><cfvo type="max"/><color rgb="FF638EC6"/></dataBar>' +
	`<extLst><ext uri="{B025F937-C7B1-47D3-B67F-A62EFF666E3E}" xmlns:x14="${X14_NS}"><x14:id>${ID_BAR.toLowerCase()}</x14:id></ext></extLst></cfRule></conditionalFormatting>` +
	'<conditionalFormatting sqref="B1:B10 D2"><cfRule type="cellIs" dxfId="0" priority="2" operator="between"><formula>1</formula><formula>10</formula></cfRule>' +
	'<cfRule type="expression" dxfId="1" priority="3" stopIfTrue="1"><formula>_xlfn.ISFORMULA(B1)</formula></cfRule></conditionalFormatting>' +
	'<conditionalFormatting sqref="C1:C10"><cfRule type="colorScale" priority="4"><colorScale><cfvo type="min"/><cfvo type="percentile" val="50"/><cfvo type="max"/>' +
	'<color rgb="FFF8696B"/><color rgb="FFFFEB84"/><color theme="6" tint="0.4"/></colorScale></cfRule></conditionalFormatting>' +
	'<conditionalFormatting sqref="D1:D10"><cfRule type="iconSet" priority="5"><iconSet iconSet="3Arrows" showValue="0" reverse="1">' +
	'<cfvo type="percent" val="0"/><cfvo type="percent" val="33"/><cfvo type="formula" val="$Z$1" gte="0"/></iconSet></cfRule></conditionalFormatting>' +
	'<conditionalFormatting sqref="E1:E10"><cfRule type="top10" dxfId="2" priority="6" rank="5" percent="1" bottom="1"/>' +
	'<cfRule type="aboveAverage" dxfId="2" priority="7" aboveAverage="0" equalAverage="1" stdDev="2"/>' +
	'<cfRule type="aboveAverage" dxfId="2" priority="8"/></conditionalFormatting>' +
	'<conditionalFormatting sqref="G1:G10"><cfRule type="containsText" dxfId="3" priority="9" operator="containsText" text="foo">' +
	'<formula>NOT(ISERROR(SEARCH("foo",G1)))</formula></cfRule>' +
	'<cfRule type="timePeriod" dxfId="3" priority="10" timePeriod="last7Days"><formula>AND(TODAY()-FLOOR(G1,1)&lt;=6,FLOOR(G1,1)&lt;=TODAY())</formula></cfRule>' +
	'<cfRule type="duplicateValues" dxfId="3" priority="11"/><cfRule type="mystery" priority="99"/></conditionalFormatting>' +
	'<conditionalFormatting pivot="1" sqref="H20"><cfRule type="containsErrors" dxfId="3" priority="12"><formula>ISERROR(H20)</formula></cfRule></conditionalFormatting>';

const EXT =
	`<extLst><ext uri="{78C0D931-6437-407d-A8EE-F0AAD7539E65}" xmlns:x14="${X14_NS}"><x14:conditionalFormattings>` +
	`<x14:conditionalFormatting xmlns:xm="${XM_NS}"><x14:cfRule type="dataBar" id="${ID_BAR}">` +
	'<x14:dataBar minLength="0" maxLength="100" gradient="0" border="1" direction="rightToLeft" negativeBarColorSameAsPositive="0" negativeBarBorderColorSameAsPositive="0" axisPosition="middle">' +
	'<x14:cfvo type="autoMin"/><x14:cfvo type="num"><xm:f>100</xm:f></x14:cfvo>' +
	'<x14:borderColor rgb="FF638EC6"/><x14:negativeFillColor rgb="FFFF0000"/><x14:negativeBorderColor rgb="FFFF0001"/><x14:axisColor rgb="FF000000"/>' +
	'</x14:dataBar></x14:cfRule><xm:sqref>A1:A10</xm:sqref></x14:conditionalFormatting>' +
	`<x14:conditionalFormatting xmlns:xm="${XM_NS}"><x14:cfRule type="iconSet" priority="13" id="${ID_STARS}">` +
	'<x14:iconSet iconSet="3Stars" custom="1"><x14:cfvo type="percent"><xm:f>0</xm:f></x14:cfvo><x14:cfvo type="percent"><xm:f>33</xm:f></x14:cfvo>' +
	'<x14:cfvo type="percent" gte="0"><xm:f>67</xm:f></x14:cfvo>' +
	'<x14:cfIcon iconSet="3Flags" iconId="0"/><x14:cfIcon iconSet="NoIcons" iconId="0"/><x14:cfIcon iconSet="3Stars" iconId="2"/></x14:iconSet>' +
	'</x14:cfRule><xm:sqref>I1:I10</xm:sqref></x14:conditionalFormatting>' +
	`<x14:conditionalFormatting xmlns:xm="${XM_NS}"><x14:cfRule type="expression" priority="14" stopIfTrue="1" id="${ID_EXPR}">` +
	'<xm:f>Lists!$A$1&gt;0</xm:f><x14:dxf><font><b/><color rgb="FF9C0006"/></font><fill><patternFill><bgColor rgb="FFFFC7CE"/></patternFill></fill></x14:dxf>' +
	'</x14:cfRule><xm:sqref>J1:J5 L1</xm:sqref></x14:conditionalFormatting>' +
	'</x14:conditionalFormattings></ext>' +
	`<ext uri="{CCE6A557-97BC-4b89-ADB6-D9C93CAAB3DF}" xmlns:x14="${X14_NS}"><x14:dataValidations count="1" xmlns:xm="${XM_NS}">` +
	'<x14:dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1"><x14:formula1><xm:f>Lists!$A$1:$A$3</xm:f></x14:formula1>' +
	'<xm:sqref>M1:M10</xm:sqref></x14:dataValidation></x14:dataValidations></ext>' +
	`<ext uri="{05C60535-1F16-4fd2-B633-F4F36F0B64E0}" xmlns:x14="${X14_NS}"><x14:sparklineGroups xmlns:xm="${XM_NS}">` +
	'<x14:sparklineGroup type="column" displayEmptyCellsAs="gap" markers="1" high="1" negative="1" lineWeight="1.5" minAxisType="custom" manualMin="0" maxAxisType="group" rightToLeft="1">' +
	'<x14:colorSeries theme="4" tint="-0.5"/><x14:colorNegative rgb="FFFF0000"/><x14:colorAxis rgb="FF000000"/><x14:colorMarkers theme="5"/>' +
	'<x14:colorFirst theme="6"/><x14:colorLast theme="7"/><x14:colorHigh rgb="FF00FF00"/><x14:colorLow theme="8"/>' +
	'<x14:sparklines><x14:sparkline><xm:f>Sheet1!A2:F2</xm:f><xm:sqref>G2</xm:sqref></x14:sparkline>' +
	'<x14:sparkline><xm:f>Sheet1!A3:F3</xm:f><xm:sqref>G3</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>' +
	'<x14:sparklineGroup><x14:sparklines><x14:sparkline><xm:f>Sheet1!A4:F4</xm:f><xm:sqref>G4</xm:sqref></x14:sparkline></x14:sparklines></x14:sparklineGroup>' +
	'</x14:sparklineGroups></ext></extLst>';

const DV_2007 =
	'<dataValidations count="4">' +
	'<dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1:A5"><formula1>"Yes,No"</formula1></dataValidation>' +
	'<dataValidation type="whole" operator="notBetween" errorStyle="warning" errorTitle="Bad" error="Enter 1-10" promptTitle="Hint" prompt="1 to 10" sqref="B1 B3:B4">' +
	'<formula1>1</formula1><formula2>10</formula2></dataValidation>' +
	'<dataValidation type="list" showDropDown="1" sqref="C1"><formula1>$D$1:$D$3</formula1></dataValidation>' +
	'<dataValidation type="custom" sqref="E1"><formula1>_xlfn.ISFORMULA(E1)</formula1></dataValidation>' +
	'<dataValidation sqref="F1" prompt="just a prompt"/>' +
	'</dataValidations>';

async function load(): Promise<Worksheet> {
	const xml = worksheet('', { after: CF_2007 + DV_2007 + EXT });
	const wb = await readWorkbook(await buildWorkbook({ sheets: [{ name: 'CF', xml }] }));
	expect(wb.warnings).toEqual([]);
	return wb.sheets[0];
}

describe('conditional formatting', () => {
	it('reads every 2007 rule type with its attributes', async () => {
		const sheet = await load();
		const f = sheet.conditionalFormats;
		expect(f[1]).toEqual({
			ranges: [
				{ top: 0, left: 1, bottom: 9, right: 1 },
				{ top: 1, left: 3, bottom: 1, right: 3 },
			],
			rules: [
				{ type: 'cellIs', priority: 2, dxfId: 0, operator: 'between', formulas: ['1', '10'] },
				{
					type: 'expression',
					priority: 3,
					dxfId: 1,
					stopIfTrue: true,
					formulas: ['ISFORMULA(B1)'],
				},
			],
		});
		expect(f[2].rules[0].colorScale).toEqual({
			cfvos: [{ type: 'min' }, { type: 'percentile', value: '50' }, { type: 'max' }],
			colors: [{ rgb: 'FFF8696B' }, { rgb: 'FFFFEB84' }, { theme: 6, tint: 0.4 }],
		});
		expect(f[3].rules[0].iconSet).toEqual({
			iconSet: '3Arrows',
			showValue: false,
			reverse: true,
			cfvos: [
				{ type: 'percent', value: '0' },
				{ type: 'percent', value: '33' },
				{ type: 'formula', value: '$Z$1', gte: false },
			],
		});
		expect(f[4].rules).toEqual([
			{ type: 'top10', priority: 6, dxfId: 2, rank: 5, percent: true, bottom: true },
			{
				type: 'aboveAverage',
				priority: 7,
				dxfId: 2,
				aboveAverage: false,
				equalAverage: true,
				stdDev: 2,
			},
			{ type: 'aboveAverage', priority: 8, dxfId: 2, aboveAverage: true },
		]);
		expect(f[5].rules).toEqual([
			{
				type: 'containsText',
				priority: 9,
				dxfId: 3,
				operator: 'containsText',
				text: 'foo',
				formulas: ['NOT(ISERROR(SEARCH("foo",G1)))'],
			},
			{
				type: 'timePeriod',
				priority: 10,
				dxfId: 3,
				timePeriod: 'last7Days',
				formulas: ['AND(TODAY()-FLOOR(G1,1)<=6,FLOOR(G1,1)<=TODAY())'],
			},
			{ type: 'duplicateValues', priority: 11, dxfId: 3 },
		]);
		expect(f[6].pivot).toBe(true);
	});

	it('merges the Excel 2010 data-bar extension into the linked 2007 rule', async () => {
		const sheet = await load();
		const rule = sheet.conditionalFormats[0].rules[0];
		expect(rule.extId).toBe(ID_BAR.toLowerCase());
		expect(rule.fromExtension).toBeUndefined();
		expect(rule.dataBar).toEqual({
			cfvos: [{ type: 'autoMin' }, { type: 'num', value: '100' }],
			color: { rgb: 'FF638EC6' },
			showValue: false,
			minLength: 0,
			maxLength: 100,
			gradient: false,
			border: true,
			direction: 'rightToLeft',
			negativeBarColorSameAsPositive: false,
			axisPosition: 'middle',
			borderColor: { rgb: 'FF638EC6' },
			negativeFillColor: { rgb: 'FFFF0000' },
			negativeBorderColor: { rgb: 'FFFF0001' },
			axisColor: { rgb: 'FF000000' },
		});
	});

	it('adds rules that exist only in the extension, with custom icons and inline dxf', async () => {
		const sheet = await load();
		const f = sheet.conditionalFormats;
		expect(f).toHaveLength(9);
		expect(f[7]).toEqual({
			ranges: [{ top: 0, left: 8, bottom: 9, right: 8 }],
			rules: [
				{
					type: 'iconSet',
					priority: 13,
					fromExtension: true,
					extId: ID_STARS,
					iconSet: {
						iconSet: '3Stars',
						cfvos: [
							{ type: 'percent', value: '0' },
							{ type: 'percent', value: '33' },
							{ type: 'percent', value: '67', gte: false },
						],
						customIcons: [
							{ iconSet: '3Flags', iconId: 0 },
							{ iconSet: 'NoIcons', iconId: 0 },
							{ iconSet: '3Stars', iconId: 2 },
						],
					},
				},
			],
		});
		expect(f[8]).toEqual({
			ranges: [
				{ top: 0, left: 9, bottom: 4, right: 9 },
				{ top: 0, left: 11, bottom: 0, right: 11 },
			],
			rules: [
				{
					type: 'expression',
					priority: 14,
					stopIfTrue: true,
					formulas: ['Lists!$A$1>0'],
					fromExtension: true,
					extId: ID_EXPR,
					dxf: {
						font: { bold: true, color: { rgb: 'FF9C0006' } },
						fill: { type: 'pattern', patternType: 'solid', bgColor: { rgb: 'FFFFC7CE' } },
					},
				},
			],
		});
	});
});

describe('data validation', () => {
	it('reads 2007 validations (showDropDown="1" hides the dropdown) and 2010 extension validations', async () => {
		const sheet = await load();
		expect(sheet.dataValidations).toEqual([
			{
				ranges: [{ top: 0, left: 0, bottom: 4, right: 0 }],
				type: 'list',
				formula1: '"Yes,No"',
				allowBlank: true,
				inCellDropdown: true,
				showInputMessage: true,
				showErrorMessage: true,
			},
			{
				ranges: [
					{ top: 0, left: 1, bottom: 0, right: 1 },
					{ top: 2, left: 1, bottom: 3, right: 1 },
				],
				type: 'whole',
				operator: 'notBetween',
				formula1: '1',
				formula2: '10',
				errorStyle: 'warning',
				errorTitle: 'Bad',
				error: 'Enter 1-10',
				promptTitle: 'Hint',
				prompt: '1 to 10',
			},
			{
				ranges: [{ top: 0, left: 2, bottom: 0, right: 2 }],
				type: 'list',
				formula1: '$D$1:$D$3',
				inCellDropdown: false,
			},
			{
				ranges: [{ top: 0, left: 4, bottom: 0, right: 4 }],
				type: 'custom',
				formula1: 'ISFORMULA(E1)',
			},
			{ ranges: [{ top: 0, left: 5, bottom: 0, right: 5 }], type: 'none', prompt: 'just a prompt' },
			{
				ranges: [{ top: 0, left: 12, bottom: 9, right: 12 }],
				type: 'list',
				formula1: 'Lists!$A$1:$A$3',
				allowBlank: true,
				inCellDropdown: true,
				showInputMessage: true,
				showErrorMessage: true,
				fromExtension: true,
			},
		]);
	});
});

describe('sparklines', () => {
	it('reads sparkline groups from the extension list', async () => {
		const sheet = await load();
		expect(sheet.sparklineGroups).toEqual([
			{
				type: 'column',
				sparklines: [
					{ formula: 'Sheet1!A2:F2', location: { row: 1, col: 6 } },
					{ formula: 'Sheet1!A3:F3', location: { row: 2, col: 6 } },
				],
				seriesColor: { theme: 4, tint: -0.5 },
				negativeColor: { rgb: 'FFFF0000' },
				markersColor: { theme: 5 },
				highColor: { rgb: 'FF00FF00' },
				lowColor: { theme: 8 },
				firstColor: { theme: 6 },
				lastColor: { theme: 7 },
				markers: true,
				high: true,
				negative: true,
				rightToLeft: true,
				lineWeight: 1.5,
				displayEmptyCellsAs: 'gap',
				minAxisType: 'custom',
				maxAxisType: 'group',
				manualMin: 0,
			},
			{ type: 'line', sparklines: [{ formula: 'Sheet1!A4:F4', location: { row: 3, col: 6 } }] },
		]);
	});
});
