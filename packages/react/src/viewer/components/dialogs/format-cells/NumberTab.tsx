import type { CellValue, FormatCategory, NegativeNumberStyle } from 'xlsx-viewer-core';
import {
	buildNumberFormat,
	CURRENCY_SYMBOL_PRESETS,
	DATE_FORMAT_PRESETS,
	formatValue,
	FRACTION_FORMAT_PRESETS,
	NEGATIVE_NUMBER_STYLES,
	SPECIAL_FORMAT_PRESETS,
	TIME_FORMAT_PRESETS,
} from 'xlsx-viewer-core';

import { useT, type Translate } from '../../../../i18n';
import { Checkbox, FIELD, FIELD_SM, GroupBox, ListBox } from '../dialog-ui';
import {
	CATEGORY_DESCRIPTIONS,
	CURRENCY_LABEL_KEYS,
	NUMBER_CATEGORIES,
	numberDraftCode,
	PRESET_LABEL_KEYS,
	SYSTEM_FORMAT_CODES,
	withCategory,
	type NumberDraft,
	type TypeListCategory,
} from './number-draft';

interface NumberTabProps {
	draft: NumberDraft;
	onChange: (next: NumberDraft) => void;
	/** Value of the active cell, shown in the Sample box. */
	sampleValue: CellValue;
	date1904: boolean;
	customCodes: readonly string[];
}

interface TypeItem {
	code: string;
	text: string;
}

/** Date and time types show the core's sample text (an asterisk marks system formats). */
const SAMPLE_TYPES: Record<'date' | 'time', readonly TypeItem[]> = {
	date: DATE_FORMAT_PRESETS.map((p) => ({
		code: p.code,
		text: `${SYSTEM_FORMAT_CODES.has(p.code) ? '*' : ''}${p.example}`,
	})),
	time: TIME_FORMAT_PRESETS.map((p) => ({
		code: p.code,
		text: `${SYSTEM_FORMAT_CODES.has(p.code) ? '*' : ''}${p.example}`,
	})),
};

/** The types a category lists; Fraction and Special types have translated names. */
function typeItems(category: TypeListCategory, t: Translate): readonly TypeItem[] {
	if (category === 'date' || category === 'time') {
		return SAMPLE_TYPES[category];
	}
	const presets = category === 'fraction' ? FRACTION_FORMAT_PRESETS : SPECIAL_FORMAT_PRESETS;
	return presets.map((p) => {
		const key = PRESET_LABEL_KEYS[p.code];
		return { code: p.code, text: key ? t(key) : p.label };
	});
}

function DecimalPlaces({
	draft,
	onChange,
}: {
	draft: NumberDraft;
	onChange: (d: NumberDraft) => void;
}) {
	const t = useT();
	return (
		<label className='flex items-center gap-2'>
			<span>{t('dialog.formatCells.decimalPlaces')}</span>
			<input
				type='number'
				min={0}
				max={30}
				className={`${FIELD_SM} w-16`}
				value={draft.decimals}
				onChange={(e) => {
					const n = Math.trunc(Number(e.target.value));
					onChange({ ...draft, decimals: Number.isFinite(n) ? Math.max(0, Math.min(30, n)) : 0 });
				}}
			/>
		</label>
	);
}

function SymbolSelect({
	draft,
	onChange,
}: {
	draft: NumberDraft;
	onChange: (d: NumberDraft) => void;
}) {
	const t = useT();
	return (
		<label className='flex items-center gap-2'>
			<span className='w-[92px]'>{t('dialog.formatCells.symbol')}</span>
			<select
				className={`${FIELD} max-w-[220px]`}
				value={draft.currencySymbol}
				onChange={(e) => onChange({ ...draft, currencySymbol: e.target.value })}
			>
				{CURRENCY_SYMBOL_PRESETS.map((p) => {
					const key = CURRENCY_LABEL_KEYS[p.id];
					return (
						<option key={p.id} value={p.id}>
							{key ? t(key) : p.label}
						</option>
					);
				})}
			</select>
		</label>
	);
}

function NegativeNumbers({
	draft,
	onChange,
}: {
	draft: NumberDraft;
	onChange: (d: NumberDraft) => void;
}) {
	const t = useT();
	const items = NEGATIVE_NUMBER_STYLES.map((n) => {
		const code = buildNumberFormat({ ...draft, negativeStyle: n.style, category: draft.category });
		const shown = formatValue(-1234.1, code);
		return { style: n.style, text: shown.text, red: Boolean(shown.color) };
	});
	return (
		<div className='flex flex-col gap-1'>
			<span>{t('dialog.formatCells.negativeNumbersLabel')}</span>
			<ListBox
				label={t('dialog.formatCells.negativeNumbers')}
				items={items}
				getKey={(i) => i.style}
				value={draft.negativeStyle}
				onChange={(key) => onChange({ ...draft, negativeStyle: key as NegativeNumberStyle })}
				render={(i, selected) => (
					<span className={i.red && !selected ? 'text-[#e00000]' : undefined}>{i.text}</span>
				)}
				className='h-[104px]'
			/>
		</div>
	);
}

function TypeList({
	draft,
	onChange,
	category,
}: {
	draft: NumberDraft;
	onChange: (d: NumberDraft) => void;
	category: TypeListCategory;
}) {
	const t = useT();
	const items = typeItems(category, t);
	return (
		<div className='flex flex-col gap-1'>
			<span>{t('dialog.formatCells.typeLabel')}</span>
			<ListBox
				label={t('dialog.formatCells.type')}
				items={items}
				getKey={(i) => i.code}
				getText={(i) => i.text}
				value={draft.types[category]}
				onChange={(code) => onChange({ ...draft, types: { ...draft.types, [category]: code } })}
				render={(i) => i.text}
				className={category === 'fraction' ? 'h-[170px]' : 'h-[150px]'}
			/>
			{category !== 'fraction' && (
				<label className='mt-1 flex flex-col gap-1'>
					<span>{t('dialog.formatCells.locale')}</span>
					<select className={FIELD} disabled value='en-US' onChange={() => undefined}>
						<option value='en-US'>{t('dialog.formatCells.localeEnUs')}</option>
					</select>
				</label>
			)}
		</div>
	);
}

function CustomEditor({
	draft,
	onChange,
	codes,
}: {
	draft: NumberDraft;
	onChange: (d: NumberDraft) => void;
	codes: readonly string[];
}) {
	const t = useT();
	return (
		<div className='flex flex-col gap-1'>
			<label className='flex flex-col gap-1'>
				<span>{t('dialog.formatCells.typeLabel')}</span>
				<input
					className={`${FIELD} font-mono text-[12px]`}
					value={draft.customCode}
					spellCheck={false}
					onChange={(e) => onChange({ ...draft, customCode: e.target.value })}
				/>
			</label>
			<ListBox
				label={t('dialog.formatCells.formatCodes')}
				items={codes}
				getKey={(c) => c}
				value={codes.includes(draft.customCode) ? draft.customCode : null}
				onChange={(c) => onChange({ ...draft, customCode: c })}
				render={(c) => <span className='font-mono text-[12px]'>{c}</span>}
				className='h-[160px]'
			/>
		</div>
	);
}

function categoryOptions(
	category: FormatCategory,
	draft: NumberDraft,
	onChange: (d: NumberDraft) => void,
	codes: readonly string[],
	t: Translate,
) {
	switch (category) {
		case 'number':
			return (
				<>
					<DecimalPlaces draft={draft} onChange={onChange} />
					<Checkbox
						label={t('dialog.formatCells.useThousands')}
						checked={draft.useThousands}
						onChange={(v) => onChange({ ...draft, useThousands: v })}
					/>
					<NegativeNumbers draft={draft} onChange={onChange} />
				</>
			);
		case 'currency':
			return (
				<>
					<DecimalPlaces draft={draft} onChange={onChange} />
					<SymbolSelect draft={draft} onChange={onChange} />
					<NegativeNumbers draft={draft} onChange={onChange} />
				</>
			);
		case 'accounting':
			return (
				<>
					<DecimalPlaces draft={draft} onChange={onChange} />
					<SymbolSelect draft={draft} onChange={onChange} />
				</>
			);
		case 'percentage':
		case 'scientific':
			return <DecimalPlaces draft={draft} onChange={onChange} />;
		case 'date':
		case 'time':
		case 'fraction':
		case 'special':
			return <TypeList draft={draft} onChange={onChange} category={category} />;
		case 'custom':
			return <CustomEditor draft={draft} onChange={onChange} codes={codes} />;
		default:
			return null;
	}
}

/** Format Cells > Number. */
export function NumberTab({ draft, onChange, sampleValue, date1904, customCodes }: NumberTabProps) {
	const t = useT();
	const code = numberDraftCode(draft);
	let sample = { text: '', color: undefined as string | undefined };
	try {
		const shown = formatValue(sampleValue as CellValue, code, { date1904 });
		sample = { text: shown.invalid ? '#####' : shown.text, color: shown.color };
	} catch {
		sample = { text: '', color: undefined };
	}
	const descriptionKey = CATEGORY_DESCRIPTIONS[draft.category];
	const description = descriptionKey ? t(descriptionKey) : '';
	const categories = NUMBER_CATEGORIES.map((c) => ({ id: c.id, label: t(c.labelKey) }));
	return (
		<div className='flex flex-col gap-3'>
			<div className='flex gap-4'>
				<div className='flex w-[132px] shrink-0 flex-col gap-1'>
					<span>{t('dialog.formatCells.categoryLabel')}</span>
					<ListBox
						label={t('dialog.formatCells.category')}
						items={categories}
						getKey={(c) => c.id}
						getText={(c) => c.label}
						value={draft.category}
						onChange={(id) => onChange(withCategory(draft, id as FormatCategory))}
						render={(c) => c.label}
						className='h-[268px]'
					/>
				</div>
				<div className='flex min-w-0 flex-1 flex-col gap-2.5'>
					<GroupBox legend={t('dialog.sample')}>
						<div
							className='h-5 truncate'
							data-testid='number-sample'
							style={{ color: sample.color }}
						>
							{sample.text}
						</div>
					</GroupBox>
					{categoryOptions(draft.category, draft, onChange, customCodes, t)}
				</div>
			</div>
			{description && (
				<p className='text-[12px] leading-snug text-muted-foreground'>{description}</p>
			)}
		</div>
	);
}
