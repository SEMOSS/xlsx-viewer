import { describeConditionalRule, listManagedRules } from 'xlsx-viewer-core';
import { ArrowDown, ArrowUp, Copy, Trash2 } from 'lucide-react';
import { useEffect, useRef, useState } from 'react';

import { useT } from '../../../i18n';
import {
	applyManagedRules,
	duplicateRule,
	formatAppliesTo,
	moveRule,
	parseAppliesTo,
	visibleRules,
	type DraftRule,
} from '../../commands/manage-rules-commands';
import { useViewerStore } from '../../state/use-store';
import { DialogButton } from '../ui/Dialog';
import { Checkbox, DialogMessage, FIELD, FormDialog, ListBox } from './dialog-ui';
import { RuleFormatSample } from './manage-rules/RuleFormatSample';
import type { DialogProps } from './types';

/** Rules that cannot stop later rules (Excel greys Stop If True for them). */
const NO_STOP = new Set(['colorScale', 'dataBar', 'iconSet']);

const TOOL =
	'inline-flex h-7 items-center gap-1.5 rounded-md border border-border bg-card px-2 text-[12px] hover:bg-accent disabled:pointer-events-none disabled:opacity-40';

/** Home > Conditional Formatting > Manage Rules: reorder, delete, duplicate and re-target rules. */
export function ManageRulesDialog({ onClose }: DialogProps) {
	const t = useT();
	const store = useViewerStore();
	const sheet = store.activeSheet;
	const selection = store.selectionOf().ranges;
	const nextId = useRef(0);
	const [rules, setRules] = useState<DraftRule[]>(() =>
		listManagedRules(sheet.conditionalFormats).map((m) => ({ ...m, id: `r${nextId.current++}` })),
	);
	const serialize = (list: readonly DraftRule[]) =>
		JSON.stringify(list.map(({ id: _id, ...m }) => m));
	// What the sheet holds now (updated by Apply); OK / Apply only write when the draft differs.
	const baseline = useRef(serialize(rules));
	const [scope, setScope] = useState<'selection' | 'sheet'>('selection');
	const shown = visibleRules(rules, scope, selection);
	const [selected, setSelected] = useState<string | null>(shown[0]?.id ?? null);
	const current = shown.find((d) => d.id === selected) ?? null;
	const [applies, setApplies] = useState(current ? formatAppliesTo(current.ranges) : '');
	const [error, setError] = useState<string | null>(null);
	const readOnly = store.readOnly;

	useEffect(() => {
		setApplies(current ? formatAppliesTo(current.ranges) : '');
		setError(null);
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [current?.id]);

	const update = (next: DraftRule[], select: string | null = selected) => {
		setRules(next);
		setSelected(select);
	};

	/** The rules with the typed "Applies to" folded in, or null (and an error) when it is invalid. */
	const withApplies = (): DraftRule[] | null => {
		if (!current || applies === formatAppliesTo(current.ranges)) {
			return rules;
		}
		const ranges = parseAppliesTo(applies, sheet.name);
		if (!ranges) {
			setError(t('message.invalidReference'));
			return null;
		}
		return rules.map((d) => (d.id === current.id ? { ...d, ranges } : d));
	};

	const commitApplies = (): boolean => {
		const next = withApplies();
		if (next && next !== rules) {
			setRules(next);
		}
		return next !== null;
	};

	const apply = (): boolean => {
		const next = withApplies();
		if (!next) {
			return false;
		}
		setRules(next);
		const text = serialize(next);
		if (text !== baseline.current && applyManagedRules(store, next)) {
			baseline.current = text;
		}
		return true;
	};
	const dirty =
		serialize(rules) !== baseline.current ||
		(current !== null && applies !== formatAppliesTo(current.ranges));

	const remove = () => {
		if (!current) {
			return;
		}
		const at = shown.findIndex((d) => d.id === current.id);
		const rest = shown.filter((d) => d.id !== current.id);
		update(
			rules.filter((d) => d.id !== current.id),
			rest[Math.min(at, rest.length - 1)]?.id ?? null,
		);
	};

	return (
		<FormDialog
			title={t('dialog.manageRules.title')}
			onClose={onClose}
			onOk={() => apply() && onClose()}
			okDisabled={readOnly}
			width={720}
			extraButtons={
				<DialogButton onClick={() => void apply()} disabled={readOnly || !dirty}>
					{t('dialog.apply')}
				</DialogButton>
			}
		>
			<label className='flex items-center gap-2'>
				<span>{t('dialog.manageRules.showRulesFor')}</span>
				<select
					className={`${FIELD} w-[200px]`}
					value={scope}
					onChange={(e) => {
						const next = e.target.value as 'selection' | 'sheet';
						setScope(next);
						setSelected(visibleRules(rules, next, selection)[0]?.id ?? null);
					}}
				>
					<option value='selection'>{t('dialog.manageRules.currentSelection')}</option>
					<option value='sheet'>{t('dialog.manageRules.thisWorksheet')}</option>
				</select>
			</label>
			<div
				className='flex flex-wrap gap-1.5'
				role='toolbar'
				aria-label={t('dialog.manageRules.rules')}
			>
				<button type='button' className={TOOL} disabled={!current || readOnly} onClick={remove}>
					<Trash2 size={14} /> {t('dialog.manageRules.deleteRule')}
				</button>
				<button
					type='button'
					className={TOOL}
					disabled={!current || readOnly}
					onClick={() => {
						if (current) {
							const id = `r${nextId.current++}`;
							update(duplicateRule(rules, current.id, id), id);
						}
					}}
				>
					<Copy size={14} /> {t('dialog.manageRules.duplicateRule')}
				</button>
				<button
					type='button'
					aria-label={t('dialog.manageRules.moveUp')}
					title={t('dialog.manageRules.moveUp')}
					className={TOOL}
					disabled={!current || readOnly || shown[0]?.id === current?.id}
					onClick={() => current && update(moveRule(rules, current.id, -1, shown))}
				>
					<ArrowUp size={14} />
				</button>
				<button
					type='button'
					aria-label={t('dialog.manageRules.moveDown')}
					title={t('dialog.manageRules.moveDown')}
					className={TOOL}
					disabled={!current || readOnly || shown[shown.length - 1]?.id === current?.id}
					onClick={() => current && update(moveRule(rules, current.id, 1, shown))}
				>
					<ArrowDown size={14} />
				</button>
			</div>
			<div className='rounded-md border border-input'>
				<div
					className='grid grid-cols-[1fr_120px_180px_84px] gap-2 border-b border-border bg-muted px-2 py-1 text-[12px] font-semibold'
					aria-hidden='true'
				>
					<span>{t('dialog.manageRules.ruleHeader')}</span>
					<span>{t('dialog.manageRules.format')}</span>
					<span>{t('dialog.manageRules.appliesTo')}</span>
					<span>{t('dialog.manageRules.stopIfTrue')}</span>
				</div>
				<ListBox
					label={t('dialog.manageRules.rules')}
					items={shown}
					getKey={(d) => d.id}
					value={current?.id ?? null}
					onChange={(id) => commitApplies() && setSelected(id)}
					render={(d) => (
						<span className='grid grid-cols-[1fr_120px_180px_84px] items-center gap-2'>
							<span className='truncate'>{describeConditionalRule(d.rule)}</span>
							<RuleFormatSample
								rule={d.rule}
								styles={store.workbook.styles}
								theme={store.workbook.theme}
							/>
							<span className='truncate'>{formatAppliesTo(d.ranges)}</span>
							<span className='text-center'>{d.rule.stopIfTrue ? t('dialog.yes') : ''}</span>
						</span>
					)}
					className='h-[188px] rounded-none border-0'
					empty={
						<div className='px-2 py-2 text-muted-foreground'>{t('dialog.manageRules.noRules')}</div>
					}
				/>
			</div>
			<div className='grid grid-cols-[1fr_auto] items-end gap-3'>
				<label className='flex flex-col gap-1'>
					<span>{t('dialog.manageRules.appliesToField')}</span>
					<input
						className={FIELD}
						value={applies}
						disabled={!current || readOnly}
						onChange={(e) => {
							setApplies(e.target.value);
							setError(null);
						}}
						onBlur={() => void commitApplies()}
					/>
				</label>
				<Checkbox
					label={t('dialog.manageRules.stopIfTrue')}
					checked={Boolean(current?.rule.stopIfTrue)}
					disabled={!current || readOnly || NO_STOP.has(current.rule.type)}
					onChange={(v) =>
						current &&
						update(
							rules.map((d) =>
								d.id === current.id ? { ...d, rule: { ...d.rule, stopIfTrue: v || undefined } } : d,
							),
						)
					}
				/>
			</div>
			{error && <DialogMessage>{error}</DialogMessage>}
		</FormDialog>
	);
}
