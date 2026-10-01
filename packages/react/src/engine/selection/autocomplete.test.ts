import { cellKey, createWorksheet } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { autoCompleteCandidate } from './autocomplete';

function column(values: Array<string | number | null>) {
	const sheet = createWorksheet('S', 1);
	values.forEach((v, r) => {
		if (v !== null) {
			sheet.cells.set(cellKey(r, 0), { value: v });
		}
	});
	return sheet;
}

describe('autoCompleteCandidate', () => {
	it('completes from the single matching entry of the list', () => {
		const sheet = column(['Region', 'North', 'South', 'East', 'West', null]);
		expect(autoCompleteCandidate(sheet, 5, 0, 'no')).toBe('North');
		expect(autoCompleteCandidate(sheet, 5, 0, 'S')).toBe('South');
		expect(autoCompleteCandidate(sheet, 5, 0, 'x')).toBeNull();
	});

	it('waits while several different entries match', () => {
		const sheet = column(['North', 'Northeast', null]);
		expect(autoCompleteCandidate(sheet, 2, 0, 'Nor')).toBeNull();
		expect(autoCompleteCandidate(sheet, 2, 0, 'Northe')).toBe('Northeast');
	});

	it('stops at blank cells and ignores numbers', () => {
		const sheet = column(['Alpha', null, 'Beta', 12, null]);
		expect(autoCompleteCandidate(sheet, 4, 0, 'Al')).toBeNull();
		expect(autoCompleteCandidate(sheet, 4, 0, 'B')).toBe('Beta');
		expect(autoCompleteCandidate(sheet, 4, 0, '1')).toBeNull();
	});

	it('treats the same entry in different case as one match', () => {
		const sheet = column(['Done', 'done', null]);
		expect(autoCompleteCandidate(sheet, 2, 0, 'd')).toMatch(/^done$/i);
	});
});
