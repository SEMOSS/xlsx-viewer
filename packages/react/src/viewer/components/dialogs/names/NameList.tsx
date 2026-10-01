import { Tag } from 'lucide-react';
import { useEffect, useId, useRef, type KeyboardEvent } from 'react';

import { useT, type TranslationKey } from '../../../../i18n';
import type { NameRow } from '../../../commands/name-commands';

function cx(...parts: Array<string | false | undefined | null>): string {
	return parts.filter(Boolean).join(' ');
}

const COLUMNS =
	'grid grid-cols-[minmax(0,1.15fr)_minmax(0,1.1fr)_minmax(0,1.45fr)_minmax(0,0.8fr)_minmax(0,1fr)]';
const GRID =
	'h-[230px] overflow-auto rounded-md border border-input bg-card text-[13px] outline-none ' +
	'focus:border-primary';
const HEADERS: readonly TranslationKey[] = [
	'dialog.names.name',
	'dialog.names.value',
	'dialog.names.refersToHeader',
	'dialog.names.scope',
	'dialog.names.comment',
];
const PAGE = 8;

export interface NameListProps {
	rows: readonly NameRow[];
	/** Key of the selected row. */
	selected: string | null;
	onSelect: (key: string) => void;
	/** Double-click on a row (Edit). */
	onActivate?: (key: string) => void;
	/** Delete key on the selected row. */
	onDelete?: () => void;
}

/**
 * The Name Manager table (`role="grid"` with a header row): Name, Value,
 * Refers To, Scope and Comment. Click or arrows select a row,
 * double-click edits it and Delete removes it.
 */
export function NameList({ rows, selected, onSelect, onActivate, onDelete }: NameListProps) {
	const t = useT();
	const id = useId();
	const gridRef = useRef<HTMLDivElement>(null);
	const index = rows.findIndex((r) => r.key === selected);
	const rowId = (i: number) => `${id}-row-${i}`;

	useEffect(() => {
		if (index >= 0) {
			gridRef.current
				?.querySelector(`[id="${rowId(index)}"]`)
				?.scrollIntoView?.({ block: 'nearest' });
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [index]);

	const move = (to: number) => {
		if (rows.length > 0) {
			onSelect(rows[Math.max(0, Math.min(rows.length - 1, to))].key);
		}
	};

	const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
		switch (e.key) {
			case 'ArrowDown':
				move(index + 1);
				break;
			case 'ArrowUp':
				move(index < 0 ? 0 : index - 1);
				break;
			case 'Home':
				move(0);
				break;
			case 'End':
				move(rows.length - 1);
				break;
			case 'PageDown':
				move(index + PAGE);
				break;
			case 'PageUp':
				move(index - PAGE);
				break;
			case 'Delete':
				onDelete?.();
				break;
			default:
				return;
		}
		e.preventDefault();
	};

	return (
		<div
			ref={gridRef}
			role='grid'
			aria-label={t('dialog.names.list')}
			aria-rowcount={rows.length + 1}
			aria-activedescendant={index >= 0 ? rowId(index) : undefined}
			tabIndex={0}
			className={GRID}
			onKeyDown={onKeyDown}
		>
			<div
				role='row'
				aria-rowindex={1}
				className={cx(
					COLUMNS,
					'sticky top-0 border-b border-border bg-muted text-[12px] font-semibold',
				)}
			>
				{HEADERS.map((key) => (
					<div key={key} role='columnheader' className='truncate px-2 py-1'>
						{t(key)}
					</div>
				))}
			</div>
			{rows.map((row, i) => {
				const isSelected = row.key === selected;
				const cells = [row.value, row.refersTo, row.scope, row.comment];
				return (
					<div
						key={row.key}
						id={rowId(i)}
						role='row'
						aria-rowindex={i + 2}
						aria-selected={isSelected}
						className={cx(
							COLUMNS,
							'cursor-default select-none',
							isSelected ? 'bg-primary text-primary-foreground' : 'hover:bg-accent',
						)}
						onClick={() => {
							gridRef.current?.focus();
							onSelect(row.key);
						}}
						onDoubleClick={() => onActivate?.(row.key)}
					>
						<div
							role='gridcell'
							className='flex min-w-0 items-center gap-1.5 px-2 py-[3px]'
							title={row.name}
						>
							<Tag size={12} aria-hidden='true' className='shrink-0 opacity-70' />
							<span className='truncate'>{row.name}</span>
						</div>
						{cells.map((text, c) => (
							<div
								key={HEADERS[c + 1]}
								role='gridcell'
								className='truncate px-2 py-[3px]'
								title={text}
							>
								{text}
							</div>
						))}
					</div>
				);
			})}
		</div>
	);
}
