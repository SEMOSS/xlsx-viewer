/**
 * Every sample on the demo's landing page opens, renders its content and
 * logs no errors.
 */
import { expect, test } from 'playwright/test';

import { cellText, collectErrors, openSample, sheetTab, type SampleId } from './support/viewer';

const SAMPLES: Array<{ id: SampleId; sheet: string; cell: [string, string]; chart?: string }> = [
	{ id: 'sales', sheet: 'Sales', cell: ['A1', 'Quarterly Sales Report'], chart: 'Sales chart' },
	{ id: 'loan', sheet: 'Loan', cell: ['A1', 'Loan Calculator'], chart: 'Balance chart' },
	{ id: 'projects', sheet: 'Tracker', cell: ['A1', 'Task'] },
	{ id: 'dynamic', sheet: 'Dynamic arrays', cell: ['E1', 'SORT by sales, descending'] },
];

for (const sample of SAMPLES) {
	test(`the ${sample.id} sample opens without errors`, async ({ page }) => {
		const errors = collectErrors(page);
		await openSample(page, sample.id);
		await expect(sheetTab(page, sample.sheet)).toHaveAttribute('aria-selected', 'true');
		await expect(cellText(page, sample.cell[0])).toHaveText(sample.cell[1]);
		if (sample.chart) {
			// Frozen panes each draw the objects they overlap, so take the first copy.
			await expect(page.getByRole('img', { name: sample.chart }).first()).toBeVisible();
		}
		expect(errors).toEqual([]);
	});
}

test('switching sheets of a sample keeps rendering', async ({ page }) => {
	const errors = collectErrors(page);
	await openSample(page, 'sales');
	for (const name of ['Formats', 'Formulas', 'Table', 'Sales']) {
		await sheetTab(page, name).click();
		await expect(page.getByRole('grid', { name: `Sheet ${name}` })).toBeVisible();
	}
	expect(errors).toEqual([]);
});
