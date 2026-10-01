import { cellKey, formatValue } from 'xlsx-viewer-core';
import {
	ChevronDown,
	DecimalsArrowLeft,
	DecimalsArrowRight,
	DollarSign,
	Percent,
} from 'lucide-react';
import { useRef, useState } from 'react';

import { useT, type TranslationKey } from '../../../../i18n';
import { changeDecimals, setNumberFormat } from '../../../commands/format-commands';
import { useActiveFormat } from '../../../hooks/useActiveFormat';
import { useViewerStore } from '../../../state/use-store';
import { useShell } from '../../../shell-context';
import { ControlCluster, RibbonButton, SplitButton } from '../../ui/buttons';
import { Popover } from '../../ui/Popover';
import { RibbonGroup } from '../RibbonGroup';

export const ACCOUNTING_FORMAT = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)';
export const COMMA_FORMAT = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)';

const PRESETS: Array<{ key: TranslationKey; code: string; category: string }> = [
	{ key: 'format.general', code: 'General', category: 'general' },
	{ key: 'format.number', code: '0.00', category: 'number' },
	{ key: 'format.currency', code: '"$"#,##0.00', category: 'currency' },
	{ key: 'format.accounting', code: ACCOUNTING_FORMAT, category: 'accounting' },
	{ key: 'format.shortDate', code: 'm/d/yyyy', category: 'date' },
	{ key: 'format.longDate', code: '[$-F800]dddd, mmmm dd, yyyy', category: 'date' },
	{ key: 'format.time', code: '[$-F400]h:mm:ss AM/PM', category: 'time' },
	{ key: 'format.percentage', code: '0.00%', category: 'percentage' },
	{ key: 'format.fraction', code: '# ?/?', category: 'fraction' },
	{ key: 'format.scientific', code: '0.00E+00', category: 'scientific' },
	{ key: 'format.text', code: '@', category: 'text' },
];

const CATEGORY_LABEL: Record<string, TranslationKey> = {
	general: 'format.general',
	number: 'format.number',
	currency: 'format.currency',
	accounting: 'format.accounting',
	date: 'format.shortDate',
	time: 'format.time',
	percentage: 'format.percentage',
	fraction: 'format.fraction',
	scientific: 'format.scientific',
	text: 'format.text',
};

function FormatSelect() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const { category } = useActiveFormat();
	const ref = useRef<HTMLButtonElement>(null);
	const [open, setOpen] = useState(false);
	const sheet = store.activeSheet;
	const { row, col } = store.selectionOf().active;
	const value = sheet.cells.get(cellKey(row, col))?.value;
	const sample =
		typeof value === 'number' ? value : typeof value === 'string' && value ? value : null;
	return (
		<>
			<button
				ref={ref}
				type='button'
				disabled={store.readOnly}
				aria-label={t('cmd.numberFormat')}
				title={t('cmd.numberFormat')}
				className='inline-flex h-7 w-[118px] items-center justify-between rounded-sm border border-input bg-card px-1.5 text-[12px] hover:border-primary disabled:opacity-40'
				onClick={() => setOpen((v) => !v)}
			>
				<span className='truncate'>{t(CATEGORY_LABEL[category] ?? 'format.general')}</span>
				<ChevronDown size={14} className='text-muted-foreground' />
			</button>
			<Popover
				open={open}
				onClose={() => setOpen(false)}
				anchorRef={ref}
				role='menu'
				ariaLabel={t('cmd.numberFormat')}
			>
				<div className='flex w-[240px] flex-col py-1'>
					{PRESETS.map((p) => (
						<button
							key={p.key}
							type='button'
							role='menuitem'
							className='flex flex-col items-start px-3 py-1 text-left hover:bg-accent'
							onClick={() => {
								setOpen(false);
								setNumberFormat(store, p.code);
							}}
						>
							<span className='text-[13px]'>{t(p.key)}</span>
							{sample !== null && (
								<span className='text-[11px] text-muted-foreground'>
									{formatValue(sample, p.code, { date1904: store.workbook.date1904 }).text}
								</span>
							)}
						</button>
					))}
					<div className='my-1 h-px bg-border' />
					<button
						type='button'
						role='menuitem'
						className='px-3 py-1.5 text-left text-[13px] hover:bg-accent'
						onClick={() => {
							setOpen(false);
							shell.openDialog('formatCells', { tab: 'number' });
						}}
					>
						{t('format.more')}
					</button>
				</div>
			</Popover>
		</>
	);
}

export function NumberGroup() {
	const t = useT();
	const store = useViewerStore();
	const readOnly = store.readOnly;
	return (
		<RibbonGroup label={t('group.number')}>
			<FormatSelect />
			<ControlCluster>
				<SplitButton
					icon={<DollarSign size={15} />}
					label={t('cmd.accounting')}
					disabled={readOnly}
					onClick={() => setNumberFormat(store, ACCOUNTING_FORMAT)}
					items={[
						{
							label: t('format.accountingDollar'),
							onSelect: () => setNumberFormat(store, ACCOUNTING_FORMAT),
						},
						{
							label: t('format.accountingPound'),
							onSelect: () =>
								setNumberFormat(
									store,
									'_-[$£-809]* #,##0.00_-;-[$£-809]* #,##0.00_-;_-[$£-809]* "-"??_-;_-@_-',
								),
						},
						{
							label: t('format.accountingEuro'),
							onSelect: () =>
								setNumberFormat(
									store,
									'_-[$€-2]\\ * #,##0.00_-;-[$€-2]\\ * #,##0.00_-;_-[$€-2]\\ * "-"??_-;_-@_-',
								),
						},
						{
							label: t('format.accountingYuan'),
							onSelect: () =>
								setNumberFormat(
									store,
									'_ [$¥-804]* #,##0.00_ ;_ [$¥-804]* -#,##0.00_ ;_ [$¥-804]* "-"??_ ;_ @_ ',
								),
						},
					]}
				/>
				<RibbonButton
					icon={<Percent size={15} />}
					label={t('cmd.percent')}
					shortcut='Ctrl+Shift+%'
					disabled={readOnly}
					onClick={() => setNumberFormat(store, '0%')}
				/>
				<RibbonButton
					icon={<span className='text-[15px] leading-none font-semibold'>,</span>}
					label={t('cmd.comma')}
					disabled={readOnly}
					onClick={() => setNumberFormat(store, COMMA_FORMAT)}
				/>
			</ControlCluster>
			<ControlCluster>
				<RibbonButton
					icon={<DecimalsArrowRight size={15} />}
					label={t('cmd.increaseDecimal')}
					disabled={readOnly}
					onClick={() => changeDecimals(store, 1)}
				/>
				<RibbonButton
					icon={<DecimalsArrowLeft size={15} />}
					label={t('cmd.decreaseDecimal')}
					disabled={readOnly}
					onClick={() => changeDecimals(store, -1)}
				/>
			</ControlCluster>
		</RibbonGroup>
	);
}
