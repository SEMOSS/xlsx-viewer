import { useRef, type KeyboardEvent } from 'react';

import { formatNumber, useLocale, useT } from '../../../../i18n';
import { PREVIEW_LINES } from './wizard-model';

export interface DataPreviewProps {
	/** Parsed fields of each line. */
	rows: readonly (readonly string[])[];
	columns: number;
	/** Step 3: column headers (data formats); clicking a column selects it. */
	headers?: readonly string[];
	selected?: number;
	onSelect?: (col: number) => void;
}

function cx(...parts: Array<string | false | undefined>): string {
	return parts.filter(Boolean).join(' ');
}

/** The wizard's "Data preview": the first lines split into columns. */
export function DataPreview({ rows, columns, headers, selected, onSelect }: DataPreviewProps) {
	const t = useT();
	const locale = useLocale();
	const buttons = useRef<Array<HTMLButtonElement | null>>([]);
	const cols = Array.from({ length: columns }, (_, i) => i);
	const pick = (col: number) => {
		if (onSelect && col >= 0 && col < columns) {
			onSelect(col);
			buttons.current[col]?.focus();
		}
	};
	const onKeyDown = (e: KeyboardEvent<HTMLTableRowElement>) => {
		const at = selected ?? 0;
		const next: Record<string, number> = {
			ArrowLeft: at - 1,
			ArrowRight: at + 1,
			Home: 0,
			End: columns - 1,
		};
		if (e.key in next) {
			e.preventDefault();
			pick(next[e.key]);
		}
	};
	const shown = (col: number) => headers !== undefined && col === selected;
	return (
		<div className='flex min-w-0 flex-col gap-1'>
			<span>{t('dialog.textToColumns.dataPreview')}</span>
			<div className='h-[150px] overflow-auto rounded-md border border-input bg-card'>
				<table
					aria-label={t('dialog.textToColumns.dataPreview')}
					className='border-collapse font-mono text-[12px]'
				>
					{headers && (
						<thead>
							<tr onKeyDown={onKeyDown}>
								{cols.map((c) => (
									<th key={c} className='p-0 text-left font-normal'>
										<button
											ref={(el) => {
												buttons.current[c] = el;
											}}
											type='button'
											aria-pressed={c === selected}
											aria-label={t('dialog.textToColumns.columnHeader', {
												n: formatNumber(c + 1, locale),
												format: headers[c],
											})}
											tabIndex={c === selected ? 0 : -1}
											className={cx(
												'w-full whitespace-nowrap border-b border-r border-border',
												'px-1.5 py-0.5 text-left',
												shown(c) ? 'bg-primary text-primary-foreground' : 'hover:bg-accent',
											)}
											onClick={() => pick(c)}
										>
											{headers[c]}
										</button>
									</th>
								))}
							</tr>
						</thead>
					)}
					<tbody>
						{rows.slice(0, PREVIEW_LINES).map((row, i) => (
							<tr key={i}>
								{cols.map((c) => (
									<td
										key={c}
										className={cx(
											'whitespace-pre border-r border-border px-1.5 align-top',
											shown(c) && 'bg-accent',
										)}
										onClick={() => pick(c)}
									>
										{row[c] ?? ''}
									</td>
								))}
							</tr>
						))}
					</tbody>
				</table>
			</div>
		</div>
	);
}
