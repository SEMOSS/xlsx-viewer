import { CalcEngine, type CalcEngineOptions } from './calc';
import type { BinaryInput } from './package/opc-package';
import { readWorkbook, type ReadOptions } from './read';
import { createWorkbook } from './sheet/workbook';
import type { Workbook } from './types/workbook';
import { writeWorkbook, type WriteOptions } from './write';

export interface CreateWorkbookOptions {
	/** Sheet names (default `['Sheet1']`). */
	sheetNames?: string[];
	/** Document title (docProps/core.xml). */
	title?: string;
}

/**
 * Public facade of the engine: load, create and save workbooks.
 *
 * ```ts
 * const handler = new XlsxHandler();
 * const workbook = await handler.load(bytes);
 * workbook.sheets[0].name; // "Sheet1"
 * const out = await handler.save(workbook); // Uint8Array of a valid .xlsx
 * ```
 */
export class XlsxHandler {
	/** Parse `.xlsx` / `.xlsm` bytes. */
	async load(data: BinaryInput, options: ReadOptions = {}): Promise<Workbook> {
		return readWorkbook(data, options);
	}

	/** Serialize a workbook to `.xlsx` bytes. Unchanged parts are copied verbatim from the source file. */
	async save(workbook: Workbook, options: WriteOptions = {}): Promise<Uint8Array> {
		return writeWorkbook(workbook, options);
	}

	/** A calculation engine bound to a workbook. */
	createCalcEngine(workbook: Workbook, options: CalcEngineOptions = {}): CalcEngine {
		const engine = new CalcEngine(workbook, options);
		engine.rebuild();
		return engine;
	}

	/** A new blank workbook (in memory) plus a handler to save it. */
	static createBlank(options: CreateWorkbookOptions = {}): {
		handler: XlsxHandler;
		workbook: Workbook;
	} {
		return { handler: new XlsxHandler(), workbook: createWorkbook(options) };
	}
}
