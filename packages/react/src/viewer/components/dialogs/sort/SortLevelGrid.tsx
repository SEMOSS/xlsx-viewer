import { useT, type TranslationKey } from '../../../../i18n';
import {
	sortOrderLabels,
	type SortKeyOption,
	type SortLevel,
} from '../../../commands/sort-dialog-commands';
import { FIELD } from '../dialog-ui';

/** A level row with a stable id (rows move up and down). */
export interface SortLevelRow extends SortLevel {
	id: number;
}

const SORT_ON: ReadonlyArray<{ value: string; labelKey: TranslationKey; enabled: boolean }> = [
	{ value: 'values', labelKey: 'dialog.sort.cellValues', enabled: true },
	{ value: 'cellColor', labelKey: 'dialog.sort.cellColor', enabled: false },
	{ value: 'fontColor', labelKey: 'dialog.sort.fontColor', enabled: false },
	{ value: 'icon', labelKey: 'dialog.sort.cfIcon', enabled: false },
];

const COLUMNS =
	'grid grid-cols-[64px_minmax(0,1.3fr)_minmax(0,1fr)_minmax(0,1.1fr)] items-center gap-2';

export interface SortLevelGridProps {
	levels: readonly SortLevelRow[];
	selected: number;
	onSelect: (index: number) => void;
	onChange: (index: number, patch: Partial<SortLevel>) => void;
	keyOptions: readonly SortKeyOption[];
	byColumns: boolean;
}

/** The "Sort by" / "Then by" rows: Column (or Row), Sort On and Order boxes per level. */
export function SortLevelGrid({
	levels,
	selected,
	onSelect,
	onChange,
	keyOptions,
	byColumns,
}: SortLevelGridProps) {
	const t = useT();
	// Sorting left to right picks rows instead of columns.
	const keyHeader = t(byColumns ? 'dialog.sort.row' : 'dialog.sort.column');
	const keyBox: TranslationKey = byColumns ? 'dialog.sort.levelRow' : 'dialog.sort.levelColumn';
	return (
		<div
			role='grid'
			aria-label={t('dialog.sort.levels')}
			className='flex h-[190px] flex-col overflow-y-auto rounded-md border border-input bg-card'
		>
			<div
				role='row'
				className={`${COLUMNS} sticky top-0 border-b border-border bg-muted px-2 py-1 text-[12px] font-semibold`}
			>
				<span role='columnheader' aria-label={t('dialog.sort.level')} />
				<span role='columnheader'>{keyHeader}</span>
				<span role='columnheader'>{t('dialog.sort.sortOn')}</span>
				<span role='columnheader'>{t('dialog.sort.order')}</span>
			</div>
			{levels.map((level, i) => {
				const option = keyOptions.find((o) => o.index === level.key);
				const [ascending, descending] = sortOrderLabels(option?.kind ?? 'text', t);
				const isSelected = i === selected;
				const n = i + 1;
				return (
					<div
						key={level.id}
						role='row'
						aria-selected={isSelected}
						className={`${COLUMNS} px-2 py-1 ${isSelected ? 'bg-accent' : ''}`}
						onClick={() => onSelect(i)}
						onFocus={() => onSelect(i)}
					>
						<span role='rowheader' className='whitespace-nowrap'>
							{t(i === 0 ? 'dialog.sort.sortBy' : 'dialog.sort.thenBy')}
						</span>
						<span role='gridcell'>
							<select
								aria-label={t(keyBox, { n })}
								className={FIELD}
								value={level.key === null ? '' : String(level.key)}
								onChange={(e) =>
									onChange(i, { key: e.target.value === '' ? null : Number(e.target.value) })
								}
							>
								{level.key === null && <option value='' />}
								{keyOptions.map((o) => (
									<option key={o.index} value={String(o.index)}>
										{o.label}
									</option>
								))}
							</select>
						</span>
						<span role='gridcell'>
							<select
								aria-label={t('dialog.sort.levelSortOn', { n })}
								className={FIELD}
								value='values'
								onChange={() => undefined}
							>
								{SORT_ON.map((s) => (
									<option key={s.value} value={s.value} disabled={!s.enabled}>
										{t(s.labelKey)}
									</option>
								))}
							</select>
						</span>
						<span role='gridcell'>
							<select
								aria-label={t('dialog.sort.levelOrder', { n })}
								className={FIELD}
								value={level.descending ? 'desc' : 'asc'}
								onChange={(e) => onChange(i, { descending: e.target.value === 'desc' })}
							>
								<option value='asc'>{ascending}</option>
								<option value='desc'>{descending}</option>
							</select>
						</span>
					</div>
				);
			})}
		</div>
	);
}
