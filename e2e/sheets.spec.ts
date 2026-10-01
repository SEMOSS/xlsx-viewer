/**
 * Sheet tabs: adding, renaming and switching sheets, and the row header
 * context menu. After each, typing goes into the active sheet's grid.
 */
import { expect, test } from 'playwright/test';

import {
	cellText,
	clickCell,
	enterValues,
	focusSink,
	grid,
	nameBox,
	openBlank,
	sheetTab,
} from './support/viewer';

test.beforeEach(async ({ page }) => {
	await openBlank(page);
});

test('adds, renames and switches sheets; typing lands in the active one', async ({ page }) => {
	await page.getByRole('button', { name: 'New sheet' }).click();
	await expect(sheetTab(page, 'Sheet2')).toHaveAttribute('aria-selected', 'true');
	await expect(grid(page, 'Sheet2')).toBeVisible();
	await expect(focusSink(page)).toBeFocused();
	await enterValues(page, 'on two');
	await expect(cellText(page, 'A1')).toHaveText('on two');

	await sheetTab(page, 'Sheet2').dblclick();
	const rename = page.getByRole('textbox', { name: 'Rename' });
	await rename.fill('Summary');
	await rename.press('Enter');
	await expect(sheetTab(page, 'Summary')).toBeVisible();
	await expect(focusSink(page)).toBeFocused();
	await enterValues(page, 'still two');
	await expect(cellText(page, 'A2')).toHaveText('still two');

	await sheetTab(page, 'Sheet1').click();
	await expect(grid(page, 'Sheet1')).toBeVisible();
	await expect(focusSink(page)).toBeFocused();
	await expect(cellText(page, 'A1')).toHaveCount(0);
	await enterValues(page, 'on one');
	await expect(cellText(page, 'A1')).toHaveText('on one');

	await sheetTab(page, 'Summary').click();
	await expect(cellText(page, 'A1')).toHaveText('on two');
});

test('Ctrl+PageDown and Ctrl+PageUp switch sheets', async ({ page }) => {
	await page.getByRole('button', { name: 'New sheet' }).click();
	await expect(grid(page, 'Sheet2')).toBeVisible();
	await page.keyboard.press('Control+PageUp');
	await expect(grid(page, 'Sheet1')).toBeVisible();
	await page.keyboard.press('Control+PageDown');
	await expect(grid(page, 'Sheet2')).toBeVisible();
});

test('row header context menu inserts a row, then typing goes to the grid', async ({ page }) => {
	await clickCell(page, 'A1');
	await enterValues(page, 'top', 'bottom');

	const header = page.locator('.xv-row-header[data-row="1"]');
	await header.click({ button: 'right' });
	await page.getByRole('menuitem', { name: 'Insert', exact: true }).click();
	await expect(page.getByRole('menu')).toHaveCount(0);

	await expect(cellText(page, 'A2')).toHaveCount(0);
	await expect(cellText(page, 'A3')).toHaveText('bottom');
	await expect(nameBox(page)).toHaveValue('A2');
	await expect(focusSink(page)).toBeFocused();
	await enterValues(page, 'middle');
	await expect(cellText(page, 'A2')).toHaveText('middle');
});
