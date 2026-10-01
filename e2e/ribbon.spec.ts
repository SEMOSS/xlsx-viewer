/**
 * Ribbon and title bar commands. Like Excel, clicking a command leaves the
 * keyboard with the grid, so the next keystroke goes into the active cell.
 */
import { expect, test } from 'playwright/test';

import {
	cellText,
	clickCell,
	enterValues,
	focusSink,
	goTo,
	openBlank,
	ribbonTab,
} from './support/viewer';

test.beforeEach(async ({ page }) => {
	await openBlank(page);
});

test('Bold follows the active cell and the grid keeps the keyboard', async ({ page }) => {
	const bold = page.getByRole('button', { name: 'Bold', exact: true });
	await clickCell(page, 'B2');
	await bold.click();
	await expect(bold).toHaveAttribute('aria-pressed', 'true');
	await expect(focusSink(page)).toBeFocused();

	await enterValues(page, 'strong');
	await expect(cellText(page, 'B2')).toHaveText('strong');
	await expect(cellText(page, 'B2').locator('span').first()).toHaveCSS('font-weight', '700');
	await expect(bold).toHaveAttribute('aria-pressed', 'false');

	await goTo(page, 'B2');
	await expect(bold).toHaveAttribute('aria-pressed', 'true');
	await page.keyboard.press('ControlOrMeta+b');
	await expect(bold).toHaveAttribute('aria-pressed', 'false');
});

test('a size typed in the Font Size box applies on Enter, then the grid has the keyboard', async ({
	page,
}) => {
	/** Rendered font size of a cell's text, in CSS pixels. */
	const fontSize = (ref: string) =>
		cellText(page, ref)
			.locator('span')
			.first()
			.evaluate((el) => Number.parseFloat(getComputedStyle(el).fontSize));

	await clickCell(page, 'B3');
	await enterValues(page, 'small', 'big');
	await goTo(page, 'B4');
	const before = await fontSize('B3');

	const size = page.getByRole('textbox', { name: 'Font Size' });
	await size.click();
	await size.fill('24');
	await size.press('Enter');
	await expect(focusSink(page)).toBeFocused();
	await expect(size).toHaveValue('24');
	// 24 pt against the default 11 pt.
	await expect.poll(() => fontSize('B4')).toBeGreaterThan(before * 2);
});

test('choosing from a dropdown menu leaves the keyboard with the grid', async ({ page }) => {
	await clickCell(page, 'C4');
	await page.getByRole('button', { name: 'Borders options' }).click();
	await page.getByRole('menuitem', { name: 'All Borders' }).click();
	await expect(page.getByRole('menu')).toHaveCount(0);
	await expect(focusSink(page)).toBeFocused();

	await enterValues(page, 'boxed');
	await expect(cellText(page, 'C4')).toHaveText('boxed');
});

test('typing in the grid closes an open dropdown', async ({ page }) => {
	await clickCell(page, 'D5');
	await page.getByRole('button', { name: 'Borders options' }).click();
	await expect(page.getByRole('menuitem', { name: 'All Borders' })).toBeVisible();

	await enterValues(page, 'typed');
	await expect(page.getByRole('menu')).toHaveCount(0);
	await expect(cellText(page, 'D5')).toHaveText('typed');
});

test('Undo and Redo in the title bar', async ({ page }) => {
	await clickCell(page, 'A1');
	await enterValues(page, 'first');
	await expect(cellText(page, 'A1')).toHaveText('first');

	await page.getByRole('button', { name: /^Undo/ }).first().click();
	await expect(cellText(page, 'A1')).toHaveCount(0);
	await page.getByRole('button', { name: /^Redo/ }).first().click();
	await expect(cellText(page, 'A1')).toHaveText('first');
	await expect(focusSink(page)).toBeFocused();
});

test('ribbon tabs switch the command set', async ({ page }) => {
	await ribbonTab(page, 'Formulas').click();
	await expect(page.getByRole('button', { name: 'Trace Precedents' })).toBeVisible();
	await ribbonTab(page, 'Home').click();
	await expect(page.getByRole('button', { name: 'Bold', exact: true })).toBeVisible();
	await expect(focusSink(page)).toBeFocused();
});
