import { useId, useState } from 'react';

import { useT } from '../../../../i18n';
import { Checkbox, FormDialog, GroupBox, Radio } from '../dialog-ui';

export interface SortOptionsValue {
	caseSensitive: boolean;
	byColumns: boolean;
}

export interface SortOptionsDialogProps extends SortOptionsValue {
	onOk: (value: SortOptionsValue) => void;
	onClose: () => void;
}

/** The Sort dialog's "Options..." box: case sensitivity and orientation. */
export function SortOptionsDialog({
	caseSensitive,
	byColumns,
	onOk,
	onClose,
}: SortOptionsDialogProps) {
	const t = useT();
	const name = useId();
	const [matchCase, setMatchCase] = useState(caseSensitive);
	const [leftToRight, setLeftToRight] = useState(byColumns);
	return (
		<FormDialog
			title={t('dialog.sort.optionsTitle')}
			onClose={onClose}
			onOk={() => onOk({ caseSensitive: matchCase, byColumns: leftToRight })}
			width={280}
		>
			<Checkbox
				label={t('dialog.sort.caseSensitive')}
				checked={matchCase}
				onChange={setMatchCase}
			/>
			<GroupBox legend={t('dialog.sort.orientation')}>
				<div className='flex flex-col gap-1.5 pt-1'>
					<Radio
						name={name}
						label={t('dialog.sort.topToBottom')}
						checked={!leftToRight}
						onChange={() => setLeftToRight(false)}
					/>
					<Radio
						name={name}
						label={t('dialog.sort.leftToRight')}
						checked={leftToRight}
						onChange={() => setLeftToRight(true)}
					/>
				</div>
			</GroupBox>
		</FormDialog>
	);
}
