/** Test helpers for the real fixtures under `src/__tests__/fixtures/real`. */
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';

import JSZip from 'jszip';

export const FIXTURE_DIR = join(__dirname, '../../__tests__/fixtures/real');

export function fixtureNames(): string[] {
	return readdirSync(FIXTURE_DIR).filter((f) => f.endsWith('.xlsx'));
}

export function fixtureBytes(name: string): Uint8Array {
	return new Uint8Array(readFileSync(join(FIXTURE_DIR, name)));
}

/** The fixture with a calculation chain added (part, workbook rel and content type). */
export async function withCalcChain(bytes: Uint8Array): Promise<Uint8Array> {
	const zip = await JSZip.loadAsync(bytes);
	zip.file(
		'xl/calcChain.xml',
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<calcChain xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><c r="A1" i="1"/></calcChain>',
	);
	const relsPath = 'xl/_rels/workbook.xml.rels';
	const rels = await zip.file(relsPath)!.async('string');
	zip.file(
		relsPath,
		rels.replace(
			'</Relationships>',
			'<Relationship Id="rId99" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" Target="calcChain.xml"/></Relationships>',
		),
	);
	const ct = await zip.file('[Content_Types].xml')!.async('string');
	zip.file(
		'[Content_Types].xml',
		ct.replace(
			'</Types>',
			'<Override PartName="/xl/calcChain.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml"/></Types>',
		),
	);
	return zip.generateAsync({ type: 'uint8array' });
}
