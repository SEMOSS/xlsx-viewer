import type {
	BorderEdge,
	BorderEdgeName,
	BorderEdgesChange,
	BorderStyle,
	ColorSpec,
	RangeBorderState,
	Theme,
} from 'xlsx-viewer-core';
import { BORDER_EDGE_NAMES } from 'xlsx-viewer-core';

import { useT, type TranslationKey } from '../../../../i18n';
import { ColorSelect } from '../ColorSelect';
import { GroupBox } from '../dialog-ui';
import { BorderPreview, EdgeIcon, StrokeLine, type EdgeDisplay } from './BorderPreview';

/** Line styles in the order of Excel's two-column Style box. */
const LINE_STYLES: readonly BorderStyle[] = [
	'none',
	'hair',
	'dotted',
	'dashDotDot',
	'dashDot',
	'dashed',
	'thin',
	'mediumDashDotDot',
	'slantDashDot',
	'mediumDashDot',
	'mediumDashed',
	'medium',
	'thick',
	'double',
];

const EDGE_LABELS: Record<BorderEdgeName, TranslationKey> = {
	top: 'dialog.formatCells.edge.top',
	bottom: 'dialog.formatCells.edge.bottom',
	left: 'dialog.formatCells.edge.left',
	right: 'dialog.formatCells.edge.right',
	insideHorizontal: 'dialog.formatCells.edge.insideHorizontal',
	insideVertical: 'dialog.formatCells.edge.insideVertical',
	diagonalUp: 'dialog.formatCells.edge.diagonalUp',
	diagonalDown: 'dialog.formatCells.edge.diagonalDown',
};

export interface BorderDraft {
	/** Edges the user changed (`null` = removed). */
	change: BorderEdgesChange;
	lineStyle: BorderStyle;
	lineColor: ColorSpec | undefined;
}

interface BorderTabProps {
	initial: RangeBorderState;
	draft: BorderDraft;
	onChange: (d: BorderDraft) => void;
	multiRow: boolean;
	multiCol: boolean;
	theme: Theme;
}

function sameLine(a: EdgeDisplay, b: BorderEdge): boolean {
	return (
		a !== null &&
		a !== 'mixed' &&
		a.style === b.style &&
		JSON.stringify(a.color ?? null) === JSON.stringify(b.color ?? null)
	);
}

/** What each edge looks like now (the user's change, else the current state). */
export function displayedEdges(
	initial: RangeBorderState,
	change: BorderEdgesChange,
): Record<BorderEdgeName, EdgeDisplay> {
	const out = {} as Record<BorderEdgeName, EdgeDisplay>;
	for (const name of BORDER_EDGE_NAMES) {
		const c = change[name];
		out[name] = c === undefined ? initial[name] : c;
	}
	return out;
}

/** Format Cells > Border. */
export function BorderTab({ initial, draft, onChange, multiRow, multiCol, theme }: BorderTabProps) {
	const t = useT();
	const shown = displayedEdges(initial, draft.change);
	const line: BorderEdge | null =
		draft.lineStyle === 'none'
			? null
			: draft.lineColor
				? { style: draft.lineStyle, color: draft.lineColor }
				: { style: draft.lineStyle };
	const available = (name: BorderEdgeName) =>
		name === 'insideHorizontal' ? multiRow : name === 'insideVertical' ? multiCol : true;
	const setEdges = (names: readonly BorderEdgeName[], value: BorderEdge | null) => {
		const change = { ...draft.change };
		for (const n of names) {
			if (available(n)) {
				change[n] = value;
			}
		}
		onChange({ ...draft, change });
	};
	const toggle = (name: BorderEdgeName) => {
		if (!available(name)) {
			return;
		}
		const current = shown[name];
		const next =
			line === null || (current !== null && current !== 'mixed' && sameLine(current, line))
				? null
				: line;
		setEdges([name], next);
	};
	const edgeButton = (name: BorderEdgeName) => (
		<button
			key={name}
			type='button'
			aria-label={t(EDGE_LABELS[name])}
			aria-pressed={shown[name] !== null && shown[name] !== 'mixed'}
			disabled={!available(name)}
			className={`flex h-7 w-7 items-center justify-center rounded-sm border ${shown[name] !== null && shown[name] !== 'mixed' ? 'border-primary bg-accent' : 'border-border hover:bg-accent'} disabled:opacity-40`}
			onClick={() => toggle(name)}
		>
			<EdgeIcon edge={name} />
		</button>
	);
	const preset = (label: string, names: readonly BorderEdgeName[] | null, disabled = false) => (
		<button
			type='button'
			disabled={disabled}
			className='flex w-[64px] flex-col items-center gap-1 rounded-sm px-1 py-1 text-[12px] hover:bg-accent disabled:opacity-40'
			onClick={() => (names ? setEdges(names, line) : setEdges(BORDER_EDGE_NAMES, null))}
		>
			<span className='flex h-9 w-9 items-center justify-center rounded-sm border border-border bg-card'>
				<svg width={24} height={24} viewBox='0 0 24 24' aria-hidden='true'>
					<rect
						x={3}
						y={3}
						width={18}
						height={18}
						className='fill-none stroke-muted-foreground'
						strokeDasharray='1 1.5'
					/>
					{names?.includes('top') && (
						<rect
							x={3}
							y={3}
							width={18}
							height={18}
							className='fill-none stroke-foreground'
							strokeWidth={1.6}
						/>
					)}
					{names?.includes('insideHorizontal') && (
						<path d='M3 12 H21 M12 3 V21' className='stroke-foreground' strokeWidth={1.6} />
					)}
				</svg>
			</span>
			{label}
		</button>
	);
	return (
		<div className='flex gap-4'>
			<GroupBox legend={t('dialog.formatCells.line')} className='w-[150px] shrink-0'>
				<span className='block pt-1'>{t('dialog.formatCells.lineStyleLabel')}</span>
				<div
					role='radiogroup'
					aria-label={t('dialog.formatCells.lineStyle')}
					className='mt-1 grid grid-cols-2 gap-x-1 gap-y-0.5 rounded-md border border-input bg-white p-1'
				>
					{LINE_STYLES.map((style) => (
						<button
							key={style}
							type='button'
							role='radio'
							aria-checked={draft.lineStyle === style}
							aria-label={style === 'none' ? t('dialog.formatCells.borderNone') : style}
							className={`flex h-[18px] items-center justify-center rounded-[2px] ${draft.lineStyle === style ? 'outline outline-1 outline-primary bg-accent' : 'hover:bg-accent'}`}
							onClick={() => onChange({ ...draft, lineStyle: style })}
						>
							{style === 'none' ? (
								<span className='text-[11px] text-[#333]'>
									{t('dialog.formatCells.borderNone')}
								</span>
							) : (
								<svg width={52} height={10} aria-hidden='true'>
									<StrokeLine x1={2} y1={5} x2={50} y2={5} edge={{ style }} theme={theme} />
								</svg>
							)}
						</button>
					))}
				</div>
				<span className='mt-2 block'>{t('dialog.formatCells.color')}</span>
				<div className='mt-1'>
					<ColorSelect
						label={t('dialog.formatCells.lineColor')}
						value={draft.lineColor}
						onChange={(c) => onChange({ ...draft, lineColor: c })}
						theme={theme}
						width={126}
					/>
				</div>
			</GroupBox>
			<div className='flex min-w-0 flex-1 flex-col gap-2'>
				<GroupBox legend={t('dialog.formatCells.presets')}>
					<div className='flex justify-center gap-3 pt-1'>
						{preset(t('dialog.formatCells.borderNone'), null)}
						{preset(t('dialog.formatCells.preset.outline'), ['top', 'bottom', 'left', 'right'])}
						{preset(
							t('dialog.formatCells.preset.inside'),
							['insideHorizontal', 'insideVertical'],
							!multiRow && !multiCol,
						)}
					</div>
				</GroupBox>
				<GroupBox legend={t('dialog.formatCells.border')}>
					<div className='flex gap-2 pt-1'>
						<div className='flex flex-col justify-between py-1'>
							{edgeButton('top')}
							{edgeButton('insideHorizontal')}
							{edgeButton('bottom')}
						</div>
						<BorderPreview
							edges={shown}
							multiRow={multiRow}
							multiCol={multiCol}
							theme={theme}
							onToggle={toggle}
						/>
					</div>
					<div className='mt-1 flex justify-between pl-9 pr-1'>
						{edgeButton('diagonalUp')}
						{edgeButton('left')}
						{edgeButton('insideVertical')}
						{edgeButton('right')}
						{edgeButton('diagonalDown')}
					</div>
				</GroupBox>
				<p className='text-[12px] leading-snug text-muted-foreground'>
					{t('dialog.formatCells.borderHint')}
				</p>
			</div>
		</div>
	);
}
