import { createContext, useContext } from 'react';

/** The password a workbook is encrypted with when saved: viewer state, not part of the model. */
export interface WorkbookPasswordState {
	/** Saves and `.xlsx` downloads are encrypted with it; `undefined` saves a plain file. */
	password: string | undefined;
	/** Set the password, or remove it with `undefined` / ''; marks the document changed. */
	setPassword: (password: string | undefined) => void;
}

export const WorkbookPasswordContext = createContext<WorkbookPasswordState | null>(null);

/** The password state of the enclosing `<ExcelViewer>` (null outside one). */
export function useWorkbookPassword(): WorkbookPasswordState | null {
	return useContext(WorkbookPasswordContext);
}
