/**
 * The viewer's UI language: switching it in the demo's picker relabels the
 * ribbon and chrome without touching the open workbook.
 */
import { expect, test } from 'playwright/test';

import { cellText, clickCell, enterValues, openBlank } from './support/viewer';

test('the language picker switches the viewer UI and keeps the workbook', async ({ page }) => {
	await openBlank(page);
	await clickCell(page, 'B2');
	await enterValues(page, 'Hallo');

	await page.getByRole('combobox', { name: 'Language' }).selectOption('de');
	await expect(page.getByRole('tab', { name: 'Einfügen', exact: true })).toBeVisible();
	await expect(page.getByRole('tab', { name: 'Insert', exact: true })).toHaveCount(0);
	await expect(page.getByRole('textbox', { name: 'Namenfeld' })).toHaveValue('B3');
	await expect(cellText(page, 'B2')).toHaveText('Hallo');

	await clickCell(page, 'C2');
	await enterValues(page, 'Welt');
	await expect(cellText(page, 'C2')).toHaveText('Welt');
});
