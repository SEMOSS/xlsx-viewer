import { ArrowDown, ArrowUp, Copy, Plus, X } from 'lucide-react';
import type { ReactNode } from 'react';

import { useT } from '../../../../i18n';
import { Checkbox } from '../dialog-ui';

function ToolButton({
	label,
	icon,
	onClick,
	disabled,
	iconOnly,
}: {
	label: string;
	icon?: ReactNode;
	onClick: () => void;
	disabled?: boolean;
	iconOnly?: boolean;
}) {
	return (
		<button
			type='button'
			aria-label={iconOnly ? label : undefined}
			title={iconOnly ? label : undefined}
			disabled={disabled}
			onClick={onClick}
			className='flex items-center gap-1 rounded-md border border-border bg-card px-2 py-1 text-[13px] hover:bg-accent disabled:opacity-50'
		>
			{icon}
			{!iconOnly && <span>{label}</span>}
		</button>
	);
}

export interface SortToolbarProps {
	canAdd: boolean;
	hasLevel: boolean;
	canMoveUp: boolean;
	canMoveDown: boolean;
	onAdd: () => void;
	onDelete: () => void;
	onCopy: () => void;
	onMove: (delta: -1 | 1) => void;
	onOptions: () => void;
	hasHeader: boolean;
	headerDisabled: boolean;
	onHeaderChange: (value: boolean) => void;
}

/** Add / Delete / Copy Level, Move Up / Down, Options... and "My data has headers". */
export function SortToolbar(p: SortToolbarProps) {
	const t = useT();
	return (
		<div className='flex flex-wrap items-center gap-1.5'>
			<ToolButton
				label={t('dialog.sort.addLevel')}
				icon={<Plus size={14} className='text-primary' />}
				onClick={p.onAdd}
				disabled={!p.canAdd}
			/>
			<ToolButton
				label={t('dialog.sort.deleteLevel')}
				icon={<X size={14} className='text-destructive' />}
				onClick={p.onDelete}
				disabled={!p.hasLevel}
			/>
			<ToolButton
				label={t('dialog.sort.copyLevel')}
				icon={<Copy size={14} />}
				onClick={p.onCopy}
				disabled={!p.hasLevel || !p.canAdd}
			/>
			<ToolButton
				label={t('dialog.sort.moveUp')}
				iconOnly
				icon={<ArrowUp size={14} />}
				onClick={() => p.onMove(-1)}
				disabled={!p.canMoveUp}
			/>
			<ToolButton
				label={t('dialog.sort.moveDown')}
				iconOnly
				icon={<ArrowDown size={14} />}
				onClick={() => p.onMove(1)}
				disabled={!p.canMoveDown}
			/>
			<ToolButton label={t('dialog.sort.options')} onClick={p.onOptions} />
			<Checkbox
				className='ml-auto'
				label={t('dialog.sort.hasHeaders')}
				checked={p.hasHeader}
				disabled={p.headerDisabled}
				onChange={p.onHeaderChange}
			/>
		</div>
	);
}
