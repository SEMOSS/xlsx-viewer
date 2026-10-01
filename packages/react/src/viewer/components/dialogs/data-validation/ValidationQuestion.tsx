import { TriangleAlert } from 'lucide-react';

import { useT, type TranslationKey } from '../../../../i18n';
import { Dialog, DialogButton } from '../../ui/Dialog';

/** Why the Data Validation dialog asks before showing its settings. */
export type ValidationQuestionKind = 'mixed' | 'partial';

const QUESTIONS: Record<ValidationQuestionKind, TranslationKey> = {
	mixed: 'dialog.dataValidation.mixedQuestion',
	partial: 'dialog.dataValidation.partialQuestion',
};

export interface ValidationQuestionProps {
	kind: ValidationQuestionKind;
	onYes: () => void;
	onNo: () => void;
	/** Cancel / Escape (closes the dialog). */
	onCancel: () => void;
}

/** Excel's Yes / No question before editing a selection with mixed validation. */
export function ValidationQuestion({ kind, onYes, onNo, onCancel }: ValidationQuestionProps) {
	const t = useT();
	return (
		<Dialog
			title={t('dialog.dataValidation.title')}
			onClose={onCancel}
			width={420}
			footer={
				<>
					<DialogButton primary onClick={onYes}>
						{t('dialog.yes')}
					</DialogButton>
					<DialogButton onClick={onNo}>{t('dialog.no')}</DialogButton>
					{kind === 'partial' && (
						<DialogButton onClick={onCancel}>{t('dialog.cancel')}</DialogButton>
					)}
				</>
			}
		>
			<div className='flex items-start gap-3'>
				<TriangleAlert
					size={32}
					strokeWidth={1.75}
					className='shrink-0 text-amber-500'
					aria-hidden='true'
				/>
				<p role='alert'>{t(QUESTIONS[kind])}</p>
			</div>
		</Dialog>
	);
}
