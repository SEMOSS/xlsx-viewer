/**
 * Opening workbooks in the demo, and the locators the specs share.
 *
 * Everything goes through what a user sees and touches: accessible names, the
 * demo's `#file-input`, the row and column headers (to click a cell where it
 * is drawn) and the `data-cell` reference each rendered cell text carries.
 *
 * @module e2e/support/viewer
 */
import { resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

import { expect, type Locator, type Page } from 'playwright/test';

const ROOT = resolve(fileURLToPath(new URL('../..', import.meta.url)));

/** Absolute path to a workbook from the core package's real-file fixtures. */
export function realFixture(name: string): string {
	return resolve(ROOT, 'packages/core/src/__tests__/fixtures/real', name);
}

/** The samples of the demo's landing page (`data-testid="sample-<id>"`). */
export type SampleId = 'sales' | 'loan' | 'projects' | 'dynamic';

/** The grid's hidden focus sink: it has the keyboard whenever no cell is being edited. */
export function focusSink(page: Page): Locator {
	return page.locator('.xv-focus-sink');
}

/** The grid of the active sheet (`role="grid"`, named "Sheet <name>"). */
export function grid(page: Page, sheet?: string): Locator {
	return sheet ? page.getByRole('grid', { name: `Sheet ${sheet}` }) : page.getByRole('grid');
}

export function nameBox(page: Page): Locator {
	return page.getByRole('textbox', { name: 'Name Box' });
}

export function formulaBar(page: Page): Locator {
	return page.getByRole('textbox', { name: 'Formula bar' });
}

/** The in-cell editor (present while a cell is edited in place). */
export function cellEditor(page: Page): Locator {
	return page.getByRole('textbox', { name: 'Cell editor' });
}

/** The text drawn for a cell, e.g. `cellText(page, 'B2')`. */
export function cellText(page: Page, ref: string): Locator {
	return page.locator(`.xv-grid [data-cell="${ref}"]`);
}

/** A sheet tab, in the "Sheets" tab list (a sheet can be named like a ribbon tab). */
export function sheetTab(page: Page, name: string): Locator {
	return page.getByRole('tablist', { name: 'Sheets' }).getByRole('tab', { name, exact: true });
}

/** A ribbon tab (Home, Insert, ...), in English. */
export function ribbonTab(page: Page, name: string): Locator {
	return page.getByRole('tablist', { name: 'Ribbon' }).getByRole('tab', { name, exact: true });
}

/** Wait until a workbook is shown and the grid has the keyboard. */
export async function expectReady(page: Page): Promise<void> {
	await expect(grid(page)).toBeVisible();
	await expect(focusSink(page)).toBeFocused();
}

export async function openBlank(page: Page): Promise<void> {
	await page.goto('/');
	await page.getByRole('button', { name: 'New workbook' }).click();
	await expectReady(page);
}

export async function openSample(page: Page, id: SampleId): Promise<void> {
	await page.goto('/');
	await page.getByTestId(`sample-${id}`).click();
	await expectReady(page);
}

/** Open a file through the demo's file input, as the file chooser would. */
export async function openFile(page: Page, path: string): Promise<void> {
	await page.goto('/');
	await page.locator('#file-input').setInputFiles(path);
	await expectReady(page);
}

function parseRef(ref: string): { row: number; col: number } {
	const m = /^([A-Z]+)(\d+)$/.exec(ref);
	if (!m) {
		throw new Error(`not a cell reference: ${ref}`);
	}
	let col = 0;
	for (const ch of m[1]) {
		col = col * 26 + (ch.charCodeAt(0) - 64);
	}
	return { row: Number(m[2]) - 1, col: col - 1 };
}

/** Where a visible cell is drawn: its column header's x and its row header's y. */
export async function cellCenter(page: Page, ref: string): Promise<{ x: number; y: number }> {
	const { row, col } = parseRef(ref);
	const colBox = await page.locator(`.xv-col-header[data-col="${col}"]`).first().boundingBox();
	const rowBox = await page.locator(`.xv-row-header[data-row="${row}"]`).first().boundingBox();
	if (!colBox || !rowBox) {
		throw new Error(`${ref} is not on screen`);
	}
	return { x: colBox.x + colBox.width / 2, y: rowBox.y + rowBox.height / 2 };
}

/** Click a visible cell with the mouse. */
export async function clickCell(
	page: Page,
	ref: string,
	options: { button?: 'left' | 'right'; clickCount?: number; modifiers?: Array<'Shift'> } = {},
): Promise<void> {
	const { x, y } = await cellCenter(page, ref);
	if (options.modifiers?.includes('Shift')) {
		await page.keyboard.down('Shift');
	}
	await page.mouse.click(x, y, { button: options.button, clickCount: options.clickCount });
	if (options.modifiers?.includes('Shift')) {
		await page.keyboard.up('Shift');
	}
}

/** Select a cell or range through the Name Box (works for cells off screen too). */
export async function goTo(page: Page, ref: string): Promise<void> {
	await nameBox(page).fill(ref);
	await nameBox(page).press('Enter');
	await expect(focusSink(page)).toBeFocused();
}

/** Type each value into the active cell and confirm it with Enter (moving down). */
export async function enterValues(page: Page, ...values: string[]): Promise<void> {
	for (const value of values) {
		await page.keyboard.type(value);
		await page.keyboard.press('Enter');
	}
}

/** Collect page errors and console errors while a spec runs. */
export function collectErrors(page: Page): string[] {
	const errors: string[] = [];
	page.on('pageerror', (e) => errors.push(e.message));
	page.on('console', (m) => {
		if (m.type() === 'error') {
			errors.push(m.text());
		}
	});
	return errors;
}
