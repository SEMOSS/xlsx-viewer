import { createColorScaleRule, createDataBarRule, createIconSetRule } from 'xlsx-viewer-core';
import { Eraser, ListOrdered } from 'lucide-react';

import { useT } from '../../../../i18n';
import { addConditionalRule, clearConditionalRules } from '../../../commands/cf-commands';
import type { ConditionalRuleKind } from '../../../commands/cf-rule-builder';
import { useViewerStore } from '../../../state/use-store';
import { useShell } from '../../../shell-context';
import { MenuButton, RibbonButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';
import { CascadeMenu, FlyoutList, type CascadeEntry, type FlyoutItem } from './CascadeMenu';
import { CellStylesButton } from './CellStylesButton';
import { FormatAsTableButton } from './FormatAsTableButton';
import { ColorScaleGallery, DataBarGallery, IconSetGallery } from './cf-galleries';

/** Ribbon icon: a small grid with highlighted cells and a data bar. */
function ConditionalFormattingIcon() {
	return (
		<svg width={15} height={15} viewBox='0 0 16 16' aria-hidden='true'>
			<rect x={1.5} y={1.5} width={13} height={13} rx={1} className='fill-none stroke-current' />
			<path
				d='M1.5 6 H14.5 M1.5 10.5 H14.5 M6 1.5 V14.5'
				className='stroke-current'
				strokeWidth={0.8}
			/>
			<rect x={6.6} y={2.3} width={7.2} height={3} fill='#E0393E' />
			<rect x={6.6} y={6.8} width={4.6} height={3} fill='#4472C4' />
			<rect x={6.6} y={11.3} width={2.4} height={2.6} fill='#70AD47' />
		</svg>
	);
}

function Glyph({ children }: { children: string }) {
	return <span className='text-[12px] font-semibold leading-none'>{children}</span>;
}

/** Home > Styles: Conditional Formatting, Format as Table, Cell Styles. */
export function StylesGroup() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const readOnly = store.readOnly;
	const ruleDialog = (rule: ConditionalRuleKind) => () =>
		shell.openDialog('conditionalRule', { rule });
	const highlight: FlyoutItem[] = [
		{
			id: 'gt',
			label: t('cf.greaterThan'),
			icon: <Glyph>&gt;</Glyph>,
			onSelect: ruleDialog('greaterThan'),
		},
		{
			id: 'lt',
			label: t('cf.lessThan'),
			icon: <Glyph>&lt;</Glyph>,
			onSelect: ruleDialog('lessThan'),
		},
		{
			id: 'between',
			label: t('cf.between'),
			icon: <Glyph>&lt;&gt;</Glyph>,
			onSelect: ruleDialog('between'),
		},
		{ id: 'eq', label: t('cf.equalTo'), icon: <Glyph>=</Glyph>, onSelect: ruleDialog('equalTo') },
		{
			id: 'text',
			label: t('cf.textContains'),
			icon: <Glyph>ab</Glyph>,
			onSelect: ruleDialog('textContains'),
		},
		{
			id: 'date',
			label: t('cf.dateOccurring'),
			icon: <Glyph>31</Glyph>,
			onSelect: ruleDialog('dateOccurring'),
		},
		{
			id: 'dup',
			label: t('cf.duplicateValues'),
			icon: <Glyph>2x</Glyph>,
			onSelect: ruleDialog('duplicateValues'),
		},
	];
	const topBottom: FlyoutItem[] = [
		{
			id: 'top10',
			label: t('cf.top10'),
			icon: <Glyph>10</Glyph>,
			onSelect: ruleDialog('top10'),
		},
		{
			id: 'topPct',
			label: t('cf.topPercent'),
			icon: <Glyph>%</Glyph>,
			onSelect: ruleDialog('topPercent'),
		},
		{
			id: 'bottom10',
			label: t('cf.bottom10'),
			icon: <Glyph>10</Glyph>,
			onSelect: ruleDialog('bottom10'),
		},
		{
			id: 'bottomPct',
			label: t('cf.bottomPercent'),
			icon: <Glyph>%</Glyph>,
			onSelect: ruleDialog('bottomPercent'),
		},
		{
			id: 'above',
			label: t('cf.aboveAverage'),
			icon: <Glyph>x&#772;+</Glyph>,
			onSelect: ruleDialog('aboveAverage'),
		},
		{
			id: 'below',
			label: t('cf.belowAverage'),
			icon: <Glyph>x&#772;-</Glyph>,
			onSelect: ruleDialog('belowAverage'),
		},
	];
	const clear: FlyoutItem[] = [
		{
			id: 'selection',
			label: t('cf.clearSelected'),
			onSelect: () => clearConditionalRules(store, 'selection'),
		},
		{
			id: 'sheet',
			label: t('cf.clearSheet'),
			onSelect: () => clearConditionalRules(store, 'sheet'),
		},
	];
	const entries: CascadeEntry[] = [
		{
			id: 'highlight',
			label: t('cf.highlightRules'),
			icon: <Glyph>&gt;</Glyph>,
			submenu: (close) => (
				<FlyoutList items={highlight} close={close} label={t('cf.highlightRules')} />
			),
		},
		{
			id: 'topBottom',
			label: t('cf.topBottomRules'),
			icon: <Glyph>10</Glyph>,
			submenu: (close) => (
				<FlyoutList items={topBottom} close={close} label={t('cf.topBottomRules')} />
			),
		},
		{ id: 'sep1', separator: true },
		{
			id: 'dataBars',
			label: t('cf.dataBars'),
			icon: <span className='h-2.5 w-3.5 bg-[linear-gradient(90deg,#638EC6,#ffffff)]' />,
			submenu: (close) => (
				<DataBarGallery
					onPick={(color, gradient) => {
						close();
						addConditionalRule(store, createDataBarRule(color, gradient));
					}}
				/>
			),
		},
		{
			id: 'colorScales',
			label: t('cf.colorScales'),
			icon: <span className='h-3 w-3 bg-[linear-gradient(180deg,#63BE7B,#FFEB84,#F8696B)]' />,
			submenu: (close) => (
				<ColorScaleGallery
					onPick={(index) => {
						close();
						addConditionalRule(store, createColorScaleRule(index));
					}}
				/>
			),
		},
		{
			id: 'iconSets',
			label: t('cf.iconSets'),
			icon: <Glyph>&#8593;</Glyph>,
			submenu: (close) => (
				<IconSetGallery
					onPick={(name) => {
						close();
						addConditionalRule(store, createIconSetRule(name));
					}}
				/>
			),
		},
		{ id: 'sep2', separator: true },
		{
			id: 'clear',
			label: t('undo.clearRules'),
			icon: <Eraser size={14} />,
			submenu: (close) => <FlyoutList items={clear} close={close} label={t('undo.clearRules')} />,
		},
		{
			id: 'manage',
			label: t('cf.manageRules'),
			icon: <ListOrdered size={14} />,
			onSelect: () => shell.openDialog('manageRules'),
		},
	];
	return (
		<RibbonGroup label={t('group.styles')}>
			<MenuButton
				icon={<ConditionalFormattingIcon />}
				label={t('cmd.conditionalFormatting')}
				disabled={readOnly}
				renderPanel={(close) => (
					<CascadeMenu entries={entries} close={close} label={t('cmd.conditionalFormatting')} />
				)}
			/>
			<FormatAsTableButton />
			<CellStylesButton />
		</RibbonGroup>
	);
}
