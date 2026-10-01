import { FUNCTION_MODULES } from './functions';
import type { FunctionDefinition } from './types';

/**
 * Lookup table of built-in (and host-registered) worksheet functions,
 * keyed by canonical upper-case name.
 */
export class FunctionRegistry {
	private readonly byName = new Map<string, FunctionDefinition>();

	constructor(definitions: Iterable<FunctionDefinition> = []) {
		for (const def of definitions) {
			this.register(def);
		}
	}

	/** Registry with every built-in function. */
	static createDefault(): FunctionRegistry {
		const reg = new FunctionRegistry();
		for (const mod of FUNCTION_MODULES) {
			for (const def of mod) {
				reg.register(def);
			}
		}
		return reg;
	}

	register(def: FunctionDefinition): void {
		this.byName.set(def.name.toUpperCase(), def);
	}

	get(name: string): FunctionDefinition | undefined {
		return this.byName.get(name.toUpperCase());
	}

	has(name: string): boolean {
		return this.byName.has(name.toUpperCase());
	}

	/** All registered names, sorted. */
	names(): string[] {
		return [...this.byName.keys()].sort();
	}

	get size(): number {
		return this.byName.size;
	}
}

let defaultRegistry: FunctionRegistry | undefined;

/** Shared default registry (created on first use). */
export function getDefaultRegistry(): FunctionRegistry {
	defaultRegistry ??= FunctionRegistry.createDefault();
	return defaultRegistry;
}
