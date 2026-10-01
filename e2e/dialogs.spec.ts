/**
 * Keyboard-opened dialogs (Go To, Find, Format Cells): each takes the
 * keyboard while open and gives it back to the grid when it closes.
 */
import { expect, test } from 'playwright/test';

import {
	cellText,
	clickCell,
	enterValues,
	focusSink,
	goTo,
	nameBox,
	openBlank,
} from './support/viewer';

test.beforeEach(async ({ page }) => {
	await openBlank(page);
});

test('Go To (Ctrl+G) jumps to a reference', async ({ page }) => {
	await page.keyboard.press('ControlOrMeta+g');
	const dialog = page.getByRole('dialog', { name: 'Go To' });
	const reference = dialog.getByRole('textbox', { name: 'Reference:' });
	await expect(reference).toBeFocused();
	await reference.fill('D10');
	await reference.press('Enter');

	await expect(dialog).toHaveCount(0);
	await expect(nameBox(page)).toHaveValue('D10');
	await expect(focusSink(page)).toBeFocused();
});

test('Find (Ctrl+F) selects the next match', async ({ page }) => {
	await clickCell(page, 'A1');
	await enterValues(page, 'apple', 'pear', 'needle', 'plum');
	await goTo(page, 'A1');

	await page.keyboard.press('ControlOrMeta+f');
	const dialog = page.getByRole('dialog', { name: 'Find and Replace' });
	const what = dialog.getByRole('textbox', { name: 'Find what' });
	await what.fill('needle');
	await what.press('Enter');
	await expect(nameBox(page)).toHaveValue('A3');

	await page.keyboard.press('Escape');
	await expect(dialog).toHaveCount(0);
	await expect(focusSink(page)).toBeFocused();
});

test('Format Cells (Ctrl+1) closes with Escape and the grid has the keyboard', async ({ page }) => {
	await clickCell(page, 'B2');
	await page.keyboard.press('ControlOrMeta+1');
	const dialog = page.getByRole('dialog', { name: 'Format Cells' });
	await expect(dialog).toBeVisible();

	await page.keyboard.press('Escape');
	await expect(dialog).toHaveCount(0);
	await expect(focusSink(page)).toBeFocused();
	await enterValues(page, 'after');
	await expect(cellText(page, 'B2')).toHaveText('after');
});
