/**
 * Files in and out: opening .xlsx and .csv files from disk, and saving a
 * workbook that then opens again with every edit intact.
 */
import { writeFile } from 'node:fs/promises';

import { expect, test } from 'playwright/test';

import {
	cellText,
	clickCell,
	enterValues,
	formulaBar,
	goTo,
	openBlank,
	openFile,
	realFixture,
	sheetTab,
} from './support/viewer';

test('opens an .xlsx file saved by Excel', async ({ page }) => {
	await openFile(page, realFixture('chart-data-fidelity--Microsoft_Excel_Worksheet.xlsx'));
	await expect(sheetTab(page, 'Sheet1')).toHaveAttribute('aria-selected', 'true');
	await expect(cellText(page, 'B1')).toHaveText('Alpha X');
	await expect(cellText(page, 'C2')).toHaveText('10');
});

test('Ctrl+S downloads an .xlsx that opens again with the edits', async ({ page }, testInfo) => {
	await openBlank(page);
	await clickCell(page, 'A1');
	await enterValues(page, 'kept', '42', '=A2*2');
	await expect(page).toHaveTitle(/^\* Book1\.xlsx/);

	const download = page.waitForEvent('download');
	await page.keyboard.press('ControlOrMeta+s');
	const file = await download;
	expect(file.suggestedFilename()).toBe('Book1.xlsx');
	const path = testInfo.outputPath('Book1.xlsx');
	await file.saveAs(path);
	await expect(page).toHaveTitle(/^Book1\.xlsx/);

	await openFile(page, path);
	await expect(cellText(page, 'A1')).toHaveText('kept');
	await expect(cellText(page, 'A3')).toHaveText('84');
	await goTo(page, 'A3');
	await expect(formulaBar(page)).toHaveValue('=A2*2');
});

test('opens a CSV file with numbers as numbers', async ({ page }, testInfo) => {
	const path = testInfo.outputPath('people.csv');
	await writeFile(path, 'name,age\nAda,36\nAlan,41\n');
	await openFile(page, path);
	await expect(cellText(page, 'A1')).toHaveText('name');
	await expect(cellText(page, 'B3')).toHaveText('41');

	await goTo(page, 'B4');
	await enterValues(page, '=SUM(B2:B3)');
	await expect(cellText(page, 'B4')).toHaveText('77');
});
