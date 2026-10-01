import type { ColorSpec, FillSpec, PatternType, Theme } from 'xlsx-viewer-core';
import { ChevronDown } from 'lucide-react';
import { useRef, useState } from 'react';

import { fillToCss } from '../../../../engine/render/fills';
import { useT, type TranslationKey } from '../../../../i18n';
import { ColorPicker } from '../../ui/ColorPicker';
import { Popover } from '../../ui/Popover';
import { DialogButton } from '../../ui/Dialog';
import { ColorSelect } from '../ColorSelect';
import { Checkbox } from '../dialog-ui';

/** Fill tab state: background color, pattern color and pattern style. */
export interface FillDraft {
	background: ColorSpec | undefined;
	patternColor: ColorSpec | undefined;
	patternStyle: PatternType;
}

/** Pattern styles as laid out in Excel's Pattern Style palette. */
const PATTERNS: ReadonlyArray<{ type: PatternType; labelKey: TranslationKey }> = [
	{ type: 'none', labelKey: 'dialog.none' },
	{ type: 'solid', labelKey: 'dialog.formatCells.pattern.solid' },
	{ type: 'darkGray', labelKey: 'dialog.formatCells.pattern.darkGray' },
	{ type: 'mediumGray', labelKey: 'dialog.formatCells.pattern.mediumGray' },
	{ type: 'lightGray', labelKey: 'dialog.formatCells.pattern.lightGray' },
	{ type: 'gray125', labelKey: 'dialog.formatCells.pattern.gray125' },
	{ type: 'gray0625', labelKey: 'dialog.formatCells.pattern.gray0625' },
	{ type: 'darkHorizontal', labelKey: 'dialog.formatCells.pattern.darkHorizontal' },
	{ type: 'darkVertical', labelKey: 'dialog.formatCells.pattern.darkVertical' },
	{ type: 'darkDown', labelKey: 'dialog.formatCells.pattern.darkDown' },
	{ type: 'darkUp', labelKey: 'dialog.formatCells.pattern.darkUp' },
	{ type: 'darkGrid', labelKey: 'dialog.formatCells.pattern.darkGrid' },
	{ type: 'darkTrellis', labelKey: 'dialog.formatCells.pattern.darkTrellis' },
	{ type: 'lightHorizontal', labelKey: 'dialog.formatCells.pattern.lightHorizontal' },
	{ type: 'lightVertical', labelKey: 'dialog.formatCells.pattern.lightVertical' },
	{ type: 'lightDown', labelKey: 'dialog.formatCells.pattern.lightDown' },
	{ type: 'lightUp', labelKey: 'dialog.formatCells.pattern.lightUp' },
	{ type: 'lightGrid', labelKey: 'dialog.formatCells.pattern.lightGrid' },
	{ type: 'lightTrellis', labelKey: 'dialog.formatCells.pattern.lightTrellis' },
];

const AUTO_FG: ColorSpec = { indexed: 64 };
const AUTO_BG: ColorSpec = { indexed: 65 };

function isAutomatic(c: ColorSpec | undefined): boolean {
	return !c || c.auto === true || c.indexed === 64 || c.indexed === 65;
}

/** The dialog's draft for an existing fill (gradients start as "no fill" until changed). */
export function fillDraftFrom(fill: FillSpec | undefined): FillDraft {
	if (!fill || fill.type !== 'pattern' || fill.patternType === 'none') {
		return { background: undefined, patternColor: undefined, patternStyle: 'none' };
	}
	if (fill.patternType === 'solid') {
		return {
			background: isAutomatic(fill.fgColor) ? undefined : fill.fgColor,
			patternColor: undefined,
			patternStyle: 'none',
		};
	}
	return {
		background: isAutomatic(fill.bgColor) ? undefined : fill.bgColor,
		patternColor: isAutomatic(fill.fgColor) ? undefined : fill.fgColor,
		patternStyle: fill.patternType,
	};
}

/** The cell fill the draft describes (`undefined` = No Color). */
export function fillFromDraft(d: FillDraft): FillSpec | undefined {
	if (d.patternStyle === 'solid') {
		return {
			type: 'pattern',
			patternType: 'solid',
			fgColor: d.patternColor ?? AUTO_FG,
			bgColor: AUTO_FG,
		};
	}
	if (d.patternStyle === 'none') {
		return d.background
			? { type: 'pattern', patternType: 'solid', fgColor: d.background, bgColor: AUTO_FG }
			: undefined;
	}
	return {
		type: 'pattern',
		patternType: d.patternStyle,
		fgColor: d.patternColor ?? AUTO_FG,
		bgColor: d.background ?? AUTO_BG,
	};
}

function Swatch({
	fill,
	theme,
	size = 18,
}: {
	fill: FillSpec | undefined;
	theme: Theme;
	size?: number;
}) {
	const css = fillToCss(fill, { theme }) ?? { backgroundColor: '#ffffff' };
	return (
		<span
			className='inline-block border border-black/20'
			style={{ width: size, height: size, ...css }}
		/>
	);
}

function PatternSelect({
	draft,
	onChange,
	theme,
}: {
	draft: FillDraft;
	onChange: (d: FillDraft) => void;
	theme: Theme;
}) {
	const t = useT();
	const ref = useRef<HTMLButtonElement>(null);
	const [open, setOpen] = useState(false);
	const sample = (type: PatternType): FillSpec | undefined =>
		type === 'none'
			? undefined
			: {
					type: 'pattern',
					patternType: type,
					fgColor: draft.patternColor ?? AUTO_FG,
					bgColor: { rgb: 'FFFFFFFF' },
				};
	const current = PATTERNS.find((p) => p.type === draft.patternStyle) ?? PATTERNS[0];
	const styleLabel = t('dialog.formatCells.patternStyle');
	return (
		<>
			<button
				ref={ref}
				type='button'
				aria-label={styleLabel}
				aria-haspopup='listbox'
				aria-expanded={open}
				className='inline-flex h-7 w-[150px] items-center gap-2 rounded-md border border-input bg-card px-1.5 text-[12px] hover:border-primary'
				onClick={() => setOpen((v) => !v)}
			>
				<Swatch fill={sample(current.type)} theme={theme} size={16} />
				<span className='flex-1 truncate text-left'>{t(current.labelKey)}</span>
				<ChevronDown size={13} className='text-muted-foreground' />
			</button>
			<Popover
				open={open}
				onClose={() => setOpen(false)}
				anchorRef={ref}
				role='listbox'
				ariaLabel={styleLabel}
			>
				<div className='grid grid-cols-6 gap-1 p-2'>
					{PATTERNS.map((p) => (
						<button
							key={p.type}
							type='button'
							role='option'
							aria-selected={p.type === draft.patternStyle}
							title={t(p.labelKey)}
							aria-label={t(p.labelKey)}
							className={`flex h-6 w-6 items-center justify-center rounded-[2px] ${p.type === draft.patternStyle ? 'outline outline-2 outline-primary' : 'hover:outline hover:outline-1 hover:outline-primary'}`}
							onClick={() => {
								setOpen(false);
								onChange({ ...draft, patternStyle: p.type });
							}}
						>
							<Swatch fill={sample(p.type)} theme={theme} />
						</button>
					))}
				</div>
			</Popover>
		</>
	);
}

/** Format Cells > Fill. */
export function FillTab({
	draft,
	onChange,
	theme,
	gradient,
}: {
	draft: FillDraft;
	onChange: (d: FillDraft) => void;
	theme: Theme;
	gradient: boolean;
}) {
	const t = useT();
	return (
		<div className='flex flex-col gap-3'>
			<div className='flex gap-6'>
				<div className='flex flex-col gap-1'>
					<span>{t('dialog.formatCells.backgroundColorLabel')}</span>
					<div
						className='rounded-md border border-border'
						role='group'
						aria-label={t('dialog.formatCells.backgroundColor')}
					>
						<ColorPicker
							theme={theme}
							noneLabel={t('dialog.formatCells.noColor')}
							close={() => undefined}
							onPick={(c) =>
								onChange(
									c
										? { ...draft, background: c }
										: { background: undefined, patternColor: undefined, patternStyle: 'none' },
								)
							}
						/>
					</div>
					<div className='mt-1 flex gap-2'>
						<DialogButton disabled>{t('dialog.formatCells.fillEffects')}</DialogButton>
					</div>
				</div>
				<div className='flex flex-col gap-1'>
					<span>{t('dialog.formatCells.patternColorLabel')}</span>
					<ColorSelect
						label={t('dialog.formatCells.patternColor')}
						value={draft.patternColor}
						onChange={(c) => onChange({ ...draft, patternColor: c })}
						theme={theme}
						width={150}
					/>
					<span className='mt-2'>{t('dialog.formatCells.patternStyleLabel')}</span>
					<PatternSelect draft={draft} onChange={onChange} theme={theme} />
					{gradient && (
						<p className='mt-2 max-w-[170px] text-[12px] text-muted-foreground'>
							{t('dialog.formatCells.gradientHint')}
						</p>
					)}
				</div>
			</div>
			<div className='flex flex-col gap-1'>
				<span>{t('dialog.sample')}</span>
				<div
					className='h-[46px] rounded-sm border border-border'
					data-testid='fill-sample'
					style={fillToCss(fillFromDraft(draft), { theme }) ?? { backgroundColor: '#ffffff' }}
				/>
			</div>
		</div>
	);
}

/** Format Cells > Protection. */
export function ProtectionTab({
	locked,
	hidden,
	onChange,
	lockedMixed,
	hiddenMixed,
}: {
	locked: boolean;
	hidden: boolean;
	lockedMixed: boolean;
	hiddenMixed: boolean;
	onChange: (p: { locked?: boolean; hidden?: boolean }) => void;
}) {
	const t = useT();
	return (
		<div className='flex flex-col gap-2'>
			<Checkbox
				label={t('dialog.formatCells.locked')}
				checked={locked}
				indeterminate={lockedMixed}
				onChange={(v) => onChange({ locked: v })}
			/>
			<Checkbox
				label={t('dialog.formatCells.hidden')}
				checked={hidden}
				indeterminate={hiddenMixed}
				onChange={(v) => onChange({ hidden: v })}
			/>
			<p className='mt-2 text-[12px] leading-snug text-muted-foreground'>
				{t('dialog.formatCells.protectionHint')}
			</p>
		</div>
	);
}
