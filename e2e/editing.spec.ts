/**
 * Entering and editing cell contents with the mouse and keyboard, as in Excel.
 *
 * The first test is the regression that motivated this suite: clicking a cell
 * used to hand keyboard focus to the page (the browser's default for a press
 * on a non-focusable element), so typing did nothing until a double-click.
 */
import { expect, test } from 'playwright/test';

import {
	cellEditor,
	cellText,
	clickCell,
	enterValues,
	focusSink,
	formulaBar,
	goTo,
	nameBox,
	openBlank,
} from './support/viewer';

test.beforeEach(async ({ page }) => {
	await openBlank(page);
});

test('typing into a clicked cell starts an edit without a double-click', async ({ page }) => {
	await clickCell(page, 'B2');
	await expect(nameBox(page)).toHaveValue('B2');
	await expect(focusSink(page)).toBeFocused();

	await page.keyboard.type('Hello');
	await expect(cellEditor(page)).toHaveValue('Hello');
	await expect(formulaBar(page)).toHaveValue('Hello');
	await page.keyboard.press('Enter');

	await expect(cellText(page, 'B2')).toHaveText('Hello');
	await expect(nameBox(page)).toHaveValue('B3');
	await expect(cellEditor(page)).toHaveCount(0);
});

test('Enter, Tab and the arrow keys move the active cell', async ({ page }) => {
	await clickCell(page, 'A1');
	await enterValues(page, 'north');
	await expect(nameBox(page)).toHaveValue('A2');

	await page.keyboard.type('west');
	await page.keyboard.press('Tab');
	await expect(nameBox(page)).toHaveValue('B2');
	await expect(cellText(page, 'A2')).toHaveText('west');

	await page.keyboard.press('ArrowRight');
	await page.keyboard.press('ArrowDown');
	await expect(nameBox(page)).toHaveValue('C3');
	await page.keyboard.press('Shift+Enter');
	await expect(nameBox(page)).toHaveValue('C2');
});

test('Escape cancels an edit and keeps the previous contents', async ({ page }) => {
	await clickCell(page, 'C3');
	await enterValues(page, 'kept');
	await clickCell(page, 'C3');
	await page.keyboard.type('discarded');
	await page.keyboard.press('Escape');

	await expect(cellText(page, 'C3')).toHaveText('kept');
	await expect(nameBox(page)).toHaveValue('C3');
	await expect(focusSink(page)).toBeFocused();
});

test('double-click and F2 edit the existing contents in place', async ({ page }) => {
	await clickCell(page, 'B4');
	await enterValues(page, 'abc');

	await clickCell(page, 'B4', { clickCount: 2 });
	await expect(cellEditor(page)).toHaveValue('abc');
	await page.keyboard.press('End');
	await page.keyboard.type('d');
	await page.keyboard.press('Enter');
	await expect(cellText(page, 'B4')).toHaveText('abcd');

	await goTo(page, 'B4');
	await page.keyboard.press('F2');
	await expect(cellEditor(page)).toHaveValue('abcd');
	await page.keyboard.type('e');
	await page.keyboard.press('Enter');
	await expect(cellText(page, 'B4')).toHaveText('abcde');
});

test('Delete clears the selection and undo brings it back', async ({ page }) => {
	await clickCell(page, 'A1');
	await enterValues(page, '1', '2', '3');
	await clickCell(page, 'A1');
	await clickCell(page, 'A3', { modifiers: ['Shift'] });
	await expect(nameBox(page)).toHaveValue('A1');

	await page.keyboard.press('Delete');
	await expect(cellText(page, 'A2')).toHaveCount(0);

	await page.keyboard.press('ControlOrMeta+z');
	await expect(cellText(page, 'A1')).toHaveText('1');
	await expect(cellText(page, 'A3')).toHaveText('3');

	await page.keyboard.press('ControlOrMeta+y');
	await expect(cellText(page, 'A1')).toHaveCount(0);
});

test('editing in the formula bar commits to the active cell', async ({ page }) => {
	await clickCell(page, 'D2');
	await formulaBar(page).click();
	await page.keyboard.type('from the bar');
	await page.keyboard.press('Enter');

	await expect(cellText(page, 'D2')).toHaveText('from the bar');
	await expect(nameBox(page)).toHaveValue('D3');
	await expect(focusSink(page)).toBeFocused();
	// The grid has the keyboard again: typing goes straight into D3.
	await enterValues(page, 'next');
	await expect(cellText(page, 'D3')).toHaveText('next');
});

test('the Name Box selects a cell and gives the keyboard back to the grid', async ({ page }) => {
	await goTo(page, 'F20');
	await expect(nameBox(page)).toHaveValue('F20');
	await enterValues(page, 'far');
	await expect(cellText(page, 'F20')).toHaveText('far');
});
