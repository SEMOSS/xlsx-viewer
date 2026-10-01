import { afterEach, describe, expect, it } from 'vitest';

import { readWorkbook } from '../read';
import { writeWorkbook } from '../write';
import { fixtureBytes } from '../write/test-utils/fixtures';
import { setNativeDeflateForTests } from './zip';

describe('package ZIP backends', () => {
	afterEach(() => setNativeDeflateForTests(undefined));

	it('reads and writes the same workbook with native streams and with the JSZip fallback', async () => {
		const bytes = fixtureBytes('ole-editable--Microsoft_Excel_Worksheet.xlsx');
		const native = await readWorkbook(bytes);
		setNativeDeflateForTests(false);
		const fallback = await readWorkbook(bytes);
		expect([...fallback.sheets[0].cells.entries()]).toEqual([...native.sheets[0].cells.entries()]);
		fallback.sheets[0].dirty = true;
		const savedByJsZip = await writeWorkbook(fallback);
		setNativeDeflateForTests(undefined);
		const back = await readWorkbook(savedByJsZip);
		expect([...back.sheets[0].cells.entries()]).toEqual([...native.sheets[0].cells.entries()]);
	});

	it('copies unchanged parts without recompressing them', async () => {
		const bytes = fixtureBytes('ole-editable--Microsoft_Excel_Worksheet.xlsx');
		const wb = await readWorkbook(bytes);
		wb.sheets[0].dirty = true;
		const saved = await writeWorkbook(wb);
		const { ZipReader } = await import('./zip');
		const before = new ZipReader(bytes);
		const after = new ZipReader(saved);
		const theme = before.entries.find((e) => e.name.includes('theme'))!;
		expect(after.raw(theme.name)?.compressed).toEqual(before.raw(theme.name)?.compressed);
	});
});
