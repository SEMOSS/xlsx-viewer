/** Props every dialog component receives from the DialogHost. */
export interface DialogProps {
	onClose: () => void;
	/** Optional data passed with `openDialog(kind, payload)`. */
	payload?: unknown;
	/** Author name for notes and comments. */
	author: string;
}
