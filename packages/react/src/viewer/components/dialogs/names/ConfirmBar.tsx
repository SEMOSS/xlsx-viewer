import { useEffect, useId, useRef } from 'react';

import { DialogButton } from '../../ui/Dialog';

export interface ConfirmRequest {
	message: string;
	okLabel: string;
	cancelLabel: string;
	onOk: () => void;
	onCancel?: () => void;
}

/**
 * A question asked inside a dialog (Excel's message box): the first button
 * takes the focus, Escape answers with the second one and does not reach
 * the enclosing dialog.
 */
export function ConfirmBar({ request, onDone }: { request: ConfirmRequest; onDone: () => void }) {
	const id = useId();
	const ref = useRef<HTMLDivElement>(null);
	useEffect(() => {
		ref.current?.querySelector<HTMLButtonElement>('button')?.focus();
	}, [request]);
	const answer = (fn?: () => void) => {
		onDone();
		fn?.();
	};
	return (
		<div
			ref={ref}
			role='alertdialog'
			aria-modal='false'
			aria-labelledby={`${id}-msg`}
			data-enter='self'
			className='flex items-center gap-2 rounded-md border border-border bg-muted px-3 py-2'
			onKeyDown={(e) => {
				if (e.key === 'Escape') {
					e.stopPropagation();
					answer(request.onCancel);
				}
			}}
		>
			<span id={`${id}-msg`} className='mr-auto'>
				{request.message}
			</span>
			<DialogButton primary onClick={() => answer(request.onOk)}>
				{request.okLabel}
			</DialogButton>
			<DialogButton onClick={() => answer(request.onCancel)}>{request.cancelLabel}</DialogButton>
		</div>
	);
}
