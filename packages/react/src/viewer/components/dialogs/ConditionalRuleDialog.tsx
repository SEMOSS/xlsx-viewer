import type { CfTimePeriod, DifferentialFormat, HighlightPresetName } from 'xlsx-viewer-core';
import { dxfBackgroundColor, HIGHLIGHT_PRESETS } from 'xlsx-viewer-core';
import { useMemo, useState, type CSSProperties, type ReactNode } from 'react';

import { useT } from '../../../i18n';
import { addConditionalRule } from '../../commands/cf-commands';
import {
	buildConditionalRule,
	CONDITIONAL_RULES,
	DATE_PERIODS,
	HIGHLIGHT_PRESET_KEYS,
	isConditionalRuleKind,
	suggestRuleValues,
	type ConditionalRuleKind,
} from '../../commands/cf-rule-builder';
import { useViewerStore } from '../../state/use-store';
import { DialogMessage, FIELD, FIELD_SM, FormDialog } from './dialog-ui';
import type { DialogProps } from './types';

const PRESETS = Object.keys(HIGHLIGHT_PRESETS) as HighlightPresetName[];

function hex(argb: string | undefined): string | undefined {
	return argb ? `#${argb.length === 8 ? argb.slice(2) : argb}` : undefined;
}

/** CSS approximation of a highlight preset for the sample swatch. */
function presetStyle(dxf: DifferentialFormat): CSSProperties {
	const fill = hex(dxfBackgroundColor(dxf.fill)?.rgb);
	const edge = dxf.border?.top;
	return {
		backgroundColor: fill ?? '#ffffff',
		color: hex(dxf.font?.color?.rgb) ?? '#000000',
		border: edge ? `1px solid ${hex(edge.color?.rgb) ?? '#000000'}` : '1px solid transparent',
	};
}

/**
 * The small Excel dialogs behind Highlight Cells Rules and Top / Bottom
 * Rules (payload `{ rule }`).
 */
export function ConditionalRuleDialog({ onClose, payload }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const requested = (payload as { rule?: unknown } | undefined)?.rule;
	const kind: ConditionalRuleKind = isConditionalRuleKind(requested) ? requested : 'greaterThan';
	const spec = CONDITIONAL_RULES[kind];
	const prompt = t(spec.promptKey);
	const initial = useMemo(() => {
		const sel = store.selectionOf();
		return suggestRuleValues(kind, store.activeSheet, sel.ranges, sel.active);
	}, [store, kind]);
	const [values, setValues] = useState<[string, string]>(initial.values);
	const [text, setText] = useState(initial.text);
	const [period, setPeriod] = useState<CfTimePeriod>('yesterday');
	const [unique, setUnique] = useState(false);
	const [rank, setRank] = useState('10');
	const [format, setFormat] = useState<HighlightPresetName>('lightRedFillDarkRedText');
	const [error, setError] = useState<string | null>(null);

	const ok = () => {
		const built = buildConditionalRule(
			kind,
			{ values, text, period, unique, rank, format },
			store.workbook.date1904,
			t,
		);
		if ('error' in built) {
			setError(built.error);
			return;
		}
		addConditionalRule(store, built.rule);
		onClose();
	};

	const valueInput = (i: 0 | 1, label: string) => (
		<input
			className={FIELD}
			aria-label={label}
			value={values[i]}
			onChange={(e) => {
				const next: [string, string] = [...values];
				next[i] = e.target.value;
				setValues(next);
				setError(null);
			}}
		/>
	);

	let input: ReactNode = null;
	switch (spec.input) {
		case 'value':
			input = valueInput(0, prompt);
			break;
		case 'between':
			input = (
				<div className='flex items-center gap-2'>
					{valueInput(0, t('dialog.cfRule.from'))}
					<span>{t('dialog.cfRule.and')}</span>
					{valueInput(1, t('dialog.cfRule.to'))}
				</div>
			);
			break;
		case 'text':
			input = (
				<input
					className={FIELD}
					aria-label={prompt}
					value={text}
					onChange={(e) => setText(e.target.value)}
				/>
			);
			break;
		case 'date':
			input = (
				<select
					className={FIELD}
					aria-label={prompt}
					value={period}
					onChange={(e) => setPeriod(e.target.value as CfTimePeriod)}
				>
					{DATE_PERIODS.map((p) => (
						<option key={p.value} value={p.value}>
							{t(p.labelKey)}
						</option>
					))}
				</select>
			);
			break;
		case 'duplicate':
			input = (
				<div className='flex items-center gap-2'>
					<select
						className={`${FIELD} w-[120px]`}
						aria-label={prompt}
						value={unique ? 'unique' : 'duplicate'}
						onChange={(e) => setUnique(e.target.value === 'unique')}
					>
						<option value='duplicate'>{t('dialog.cfRule.duplicate')}</option>
						<option value='unique'>{t('dialog.cfRule.unique')}</option>
					</select>
					<span>{t('dialog.cfRule.values')}</span>
				</div>
			);
			break;
		case 'rank':
			input = (
				<div className='flex items-center gap-2'>
					<input
						type='number'
						min={1}
						max={spec.percent ? 100 : 1000}
						className={`${FIELD_SM} w-20`}
						aria-label={prompt}
						value={rank}
						onChange={(e) => {
							setRank(e.target.value);
							setError(null);
						}}
					/>
					{spec.percent && <span>%</span>}
				</div>
			);
			break;
		default:
			input = <span>{t('dialog.cfRule.forSelectedRange')}</span>;
	}

	return (
		<FormDialog
			title={t(spec.titleKey)}
			onClose={onClose}
			onOk={ok}
			okDisabled={store.readOnly}
			width={spec.input === 'between' ? 560 : 480}
		>
			<span>{prompt}</span>
			<div className='flex items-center gap-2'>
				<div className='min-w-0 flex-1'>{input}</div>
				<span>{t('dialog.cfRule.with')}</span>
				<select
					className={`${FIELD} w-[220px]`}
					aria-label={t('dialog.cfRule.formatWith')}
					value={format}
					onChange={(e) => setFormat(e.target.value as HighlightPresetName)}
				>
					{PRESETS.map((p) => (
						<option key={p} value={p}>
							{t(HIGHLIGHT_PRESET_KEYS[p])}
						</option>
					))}
				</select>
				<span
					className='shrink-0 rounded-[2px] px-2 py-0.5 text-[12px]'
					style={presetStyle(HIGHLIGHT_PRESETS[format])}
					data-testid='rule-format-sample'
				>
					AaBbCc
				</span>
			</div>
			{error && <DialogMessage>{error}</DialogMessage>}
		</FormDialog>
	);
}
