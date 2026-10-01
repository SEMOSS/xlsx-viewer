/**
 * Formulas typed into the grid: calculation, recalculation, pointing at cells
 * while typing a formula, function autocomplete and error values.
 */
import { expect, test } from 'playwright/test';

import {
	cellEditor,
	cellText,
	clickCell,
	enterValues,
	formulaBar,
	goTo,
	openBlank,
} from './support/viewer';

test.beforeEach(async ({ page }) => {
	await openBlank(page);
});

test('formulas calculate and follow their precedents', async ({ page }) => {
	await clickCell(page, 'A1');
	await enterValues(page, '10', '20', '30', '=SUM(A1:A3)');
	await expect(cellText(page, 'A4')).toHaveText('60');

	await goTo(page, 'A2');
	await enterValues(page, '25');
	await expect(cellText(page, 'A4')).toHaveText('65');

	await goTo(page, 'A4');
	await expect(formulaBar(page)).toHaveValue('=SUM(A1:A3)');
});

test('clicking cells while typing a formula inserts their references', async ({ page }) => {
	await clickCell(page, 'A1');
	await enterValues(page, '4', '5');

	await clickCell(page, 'C1');
	await page.keyboard.type('=');
	await clickCell(page, 'A1');
	await page.keyboard.type('*');
	await clickCell(page, 'A2');
	await expect(cellEditor(page)).toHaveValue('=A1*A2');
	await page.keyboard.press('Enter');

	await expect(cellText(page, 'C1')).toHaveText('20');
	await goTo(page, 'C1');
	await expect(formulaBar(page)).toHaveValue('=A1*A2');
});

test('function names autocomplete while typing', async ({ page }) => {
	await clickCell(page, 'B2');
	await page.keyboard.type('=AVERA');
	const first = page
		.getByRole('listbox', { name: 'Formula suggestions' })
		.getByRole('option')
		.first();
	await expect(first).toContainText('AVERAGE');
	await expect(first).toHaveAttribute('aria-selected', 'true');

	await page.keyboard.press('Tab');
	await expect(cellEditor(page)).toHaveValue('=AVERAGE(');
	await page.keyboard.type('2,4)');
	await page.keyboard.press('Enter');
	await expect(cellText(page, 'B2')).toHaveText('3');
});

test('errors show as Excel error values', async ({ page }) => {
	await clickCell(page, 'A1');
	await enterValues(page, '=1/0', '=NOSUCHFUNCTION(1)', '=A1+1');
	await expect(cellText(page, 'A1')).toHaveText('#DIV/0!');
	await expect(cellText(page, 'A2')).toHaveText('#NAME?');
	await expect(cellText(page, 'A3')).toHaveText('#DIV/0!');
});

test('AutoSum sums the numbers above the active cell', async ({ page }) => {
	await clickCell(page, 'B1');
	await enterValues(page, '1.5', '2.5', '6');
	await page.keyboard.press('Alt+=');
	await expect(cellEditor(page)).toHaveValue('=SUM(B1:B3)');
	await page.keyboard.press('Enter');
	await expect(cellText(page, 'B4')).toHaveText('10');
});
