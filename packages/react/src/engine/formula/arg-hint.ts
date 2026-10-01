import type { FunctionInfo } from 'xlsx-viewer-core';

export interface ArgHintPart {
	label: string;
	current: boolean;
	/** Render `, ...` after this part (end of a repeating group). */
	ellipsis?: boolean;
}

export interface ArgHint {
	name: string;
	parts: ArgHintPart[];
}

/** `number2` -> `number5` for the third extra copy of a repeating group. */
function bump(name: string, step: number): string {
	return name.replace(/(\d+)$/, (d) => String(Number(d) + step));
}

/**
 * Argument tooltip content in Excel's style, `SUM(number1, [number2], [number3], ...)`,
 * with the argument at `argIndex` marked current. A repeating group (consecutive
 * repeating arguments, e.g. SUMIFS' `criteria_range2, criteria2`) is extended so
 * the current argument is always shown; arguments after a group (LAMBDA's
 * `calculation`) stay last.
 */
export function buildArgHint(info: FunctionInfo, argIndex: number): ArgHint {
	const args = info.args;
	const firstRep = args.findIndex((a) => a.repeating);
	const label = (name: string, optional?: boolean) => (optional ? `[${name}]` : name);
	if (firstRep < 0) {
		return {
			name: info.name,
			parts: args.map((a, i) => ({ label: label(a.name, a.optional), current: i === argIndex })),
		};
	}
	let groupEnd = firstRep;
	while (groupEnd < args.length && args[groupEnd].repeating) {
		groupEnd++;
	}
	const group = groupEnd - firstRep;
	const parts: ArgHintPart[] = args
		.slice(0, groupEnd)
		.map((a, i) => ({ label: label(a.name, a.optional), current: i === argIndex }));
	// Extend whole copies of the group until the current argument is covered.
	for (let p = groupEnd; p <= argIndex || (p - firstRep) % group !== 0; p++) {
		const template = args[firstRep + ((p - firstRep) % group)];
		parts.push({
			label: `[${bump(template.name, Math.floor((p - firstRep) / group))}]`,
			current: p === argIndex,
		});
	}
	parts[parts.length - 1] = { ...parts[parts.length - 1], ellipsis: true };
	// Arguments after the group: the caret index cannot tell them from more group copies.
	for (const a of args.slice(groupEnd)) {
		parts.push({ label: label(a.name, a.optional), current: false });
	}
	return { name: info.name, parts };
}
