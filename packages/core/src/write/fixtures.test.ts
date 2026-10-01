import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { OpcPackage } from '../package';
import { writeWorkbook } from './index';
import { cleanModel } from './test-utils/clean-model';
import { openPackage, validatePackage } from './test-utils/package-check';

const FIXTURES = join(__dirname, '../__tests__/fixtures/real');
const files = readdirSync(FIXTURES).filter((f) => f.endsWith('.xlsx'));

async function load(file: string): Promise<OpcPackage> {
	return OpcPackage.open(readFileSync(join(FIXTURES, file)));
}

describe('real fixtures: unchanged model', () => {
	it('has fixtures', () => {
		expect(files.length).toBeGreaterThan(5);
	});

	it.each(files)('%s keeps every part byte-identical', async (file) => {
		const src = await load(file);
		const out = await openPackage(await writeWorkbook(await cleanModel(src)));
		expect(out.partPaths().sort()).toEqual(src.partPaths().sort());
		for (const p of src.partPaths()) {
			if (p === '[Content_Types].xml') {
				continue;
			}
			expect(await out.bytes(p), p).toEqual(await src.bytes(p));
		}
		expect([...out.contentTypes.overrides].sort()).toEqual([...src.contentTypes.overrides].sort());
		expect([...out.contentTypes.defaults].sort()).toEqual([...src.contentTypes.defaults].sort());
	});

	it.each(files)('%s validates after an unchanged save', async (file) => {
		const out = await openPackage(await writeWorkbook(await cleanModel(await load(file))));
		expect(await validatePackage(out)).toEqual([]);
	});
});

describe('real fixtures: forced regeneration', () => {
	it.each(files)('%s stays valid with regenerateAll', async (file) => {
		const model = await cleanModel(await load(file));
		const out = await openPackage(await writeWorkbook(model, { regenerateAll: true }));
		expect(await validatePackage(out)).toEqual([]);
	});
});
