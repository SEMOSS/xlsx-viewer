import { describe, expect, it } from 'vitest';

import { addFormulaPrefixes, stripFormulaPrefixes } from '../prefixes';
import { translateFormula } from '../translate';
import { adjustFormulaForStructureChange } from './adjust';
import { parseFormula } from './parser';
import { serializeFormula } from './serializer';
import { tokenizeFormula } from './tokenizer';

const TEMPLATES: Array<(r: number) => string> = [
	(r) => `SUM(A${r}:A${r + 20})*2`,
	(r) => `IF(B${r}>0,VLOOKUP(B${r},Sheet2!$A$1:$D$500,3,FALSE),"")`,
	(r) => `A${r}*$B$1+C${r}`,
	(r) => `IFERROR(INDEX(Data!$C:$C,MATCH($A${r},Data!$A:$A,0)),0)`,
	(r) => `ROUND(AVERAGE(D${r}:H${r}),2)`,
	(r) => `"Total: "&TEXT(SUM(Table1[Amount]),"#,##0.00")`,
	(r) =>
		`SUMIFS('Sales Data'!$E:$E,'Sales Data'!$B:$B,$A${r},'Sales Data'!$C:$C,">="&DATE(2024,1,1))`,
	(r) => `XLOOKUP(A${r},B:B,C:C,"n/a")+LET(x,A${r}*2,x^2-1)`,
	(r) => `AND(A${r}<>"",OR(B${r}=1,B${r}=2),NOT(ISBLANK(C${r})))`,
	(r) => `[@Qty]*[@Price]*(1-[@Discount])+${r}%`,
];

function corpus(n: number): string[] {
	const out: string[] = [];
	for (let i = 0; i < n; i++) {
		out.push(TEMPLATES[i % TEMPLATES.length]((i % 5000) + 1));
	}
	return out;
}

describe('performance sanity', () => {
	it('tokenizes and parses 100k typical formulas quickly', () => {
		const formulas = corpus(100_000);
		parseFormula(formulas[0]); // warm-up
		const t0 = performance.now();
		let nodes = 0;
		for (const f of formulas) {
			nodes += parseFormula(f).type === 'binary' ? 1 : 0;
		}
		const elapsed = performance.now() - t0;
		expect(nodes).toBeGreaterThan(0);
		// Typically well under a second; the bound is generous to stay reliable on slow CI machines.
		expect(elapsed).toBeLessThan(6000);
	});

	it('rewrites formulas quickly', () => {
		const formulas = corpus(20_000);
		const t0 = performance.now();
		for (const f of formulas) {
			tokenizeFormula(f);
			translateFormula(f, 1, 1);
			adjustFormulaForStructureChange(
				f,
				{ kind: 'insertRows', sheet: 'Sheet1', at: 10, count: 2 },
				'Sheet1',
			);
			addFormulaPrefixes(stripFormulaPrefixes(f));
			serializeFormula(parseFormula(f), { forFile: true });
		}
		expect(performance.now() - t0).toBeLessThan(6000);
	});
});
