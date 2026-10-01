import { readFileSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

import { describe, expect, it } from 'vitest';

import { REL } from '../xml/namespaces';
import { OpcPackage, PackageWriter, isZipFile } from './opc-package';
import { relativeTarget, resolveRelTarget } from './paths';
import { findRelByType } from './relationships';

const FIXTURES = join(__dirname, '../__tests__/fixtures/real');

describe('part paths', () => {
	it('resolves relative and absolute targets', () => {
		expect(resolveRelTarget('xl/workbook.xml', 'worksheets/sheet1.xml')).toBe(
			'xl/worksheets/sheet1.xml',
		);
		expect(resolveRelTarget('xl/worksheets/sheet1.xml', '../drawings/drawing1.xml')).toBe(
			'xl/drawings/drawing1.xml',
		);
		expect(resolveRelTarget('xl/workbook.xml', '/xl/styles.xml')).toBe('xl/styles.xml');
		expect(relativeTarget('xl/worksheets/sheet1.xml', 'xl/drawings/drawing1.xml')).toBe(
			'../drawings/drawing1.xml',
		);
		expect(relativeTarget('xl/workbook.xml', 'xl/worksheets/sheet2.xml')).toBe(
			'worksheets/sheet2.xml',
		);
	});
});

describe('OpcPackage', () => {
	const files = readdirSync(FIXTURES).filter((f) => f.endsWith('.xlsx'));

	it('opens every real fixture and finds the workbook part', async () => {
		expect(files.length).toBeGreaterThan(5);
		for (const f of files) {
			const bytes = readFileSync(join(FIXTURES, f));
			expect(isZipFile(bytes)).toBe(true);
			const pkg = await OpcPackage.open(bytes);
			const rels = await pkg.rels('');
			const main = findRelByType(rels, REL.officeDocument);
			expect(main, f).toBeDefined();
			const wbPath = pkg.resolve('', main!);
			expect(await pkg.text(wbPath)).toContain('<sheets>');
		}
	});

	it('re-packages untouched parts byte-for-byte', async () => {
		const bytes = readFileSync(join(FIXTURES, files[0]));
		const src = await OpcPackage.open(bytes);
		const writer = new PackageWriter(src);
		writer.setText('xl/custom.xml', '<x/>', 'application/xml');
		const out = await writer.generate();
		const reread = await OpcPackage.open(out);
		for (const p of src.partPaths()) {
			if (p === '[Content_Types].xml') {
				continue;
			}
			expect(await reread.bytes(p), p).toEqual(await src.bytes(p));
		}
		expect(await reread.text('xl/custom.xml')).toBe('<x/>');
		expect(reread.contentTypes.get('xl/custom.xml')).toBe('application/xml');
	});
});
