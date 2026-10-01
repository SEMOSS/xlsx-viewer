import type { ChartElementOptions } from 'xlsx-viewer-core';
import { CHART_QUICK_LAYOUT_COUNT, chartElementState, chartTypeTargetOf } from 'xlsx-viewer-core';
import {
	ArrowLeftRight,
	ChartColumn,
	LayoutDashboard,
	Paintbrush,
	Palette,
	SquarePlus,
} from 'lucide-react';

import { useT, type TranslationKey } from '../../../../i18n';
import {
	applySelectedChartLayout,
	applySelectedChartPalette,
	applySelectedChartStyle,
	changeChartType,
	selectedChart,
	setSelectedChartElement,
	switchSelectedChartRowColumn,
} from '../../../commands/chart-commands';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import { MenuButton, RibbonButton } from '../../ui/buttons';
import type { MenuItem } from '../../ui/Menu';
import { RibbonGroup } from '../RibbonGroup';
import { ChartColorsGallery, ChartStylesGallery, ChartTypeGallery } from './chart-galleries';
import { ChartTypeGlyph } from './chart-glyphs';

type Choice<K extends keyof ChartElementOptions> = Array<[ChartElementOptions[K], TranslationKey]>;

const TITLE_CHOICES: Choice<'chartTitle'> = [
	['none', 'chart.option.none'],
	['above', 'chart.option.above'],
	['overlay', 'chart.option.overlay'],
];
const LEGEND_CHOICES: Choice<'legend'> = [
	['none', 'chart.option.none'],
	['right', 'chart.option.right'],
	['top', 'chart.option.top'],
	['left', 'chart.option.left'],
	['bottom', 'chart.option.bottom'],
];
const LABEL_CHOICES: Choice<'dataLabels'> = [
	['none', 'chart.option.none'],
	['show', 'chart.option.show'],
];

/** Contextual ribbon tab shown while a chart is selected. */
export function ChartDesignTab() {
	const t = useT();
	const store = useViewerStore();
	// The drawing object is replaced on every chart edit, so identity tracks changes.
	const object = useStoreSelector(
		(snap, st) => {
			void snap.docVersion;
			void snap.ui;
			return selectedChart(st);
		},
		(a, b) => a === b,
	);
	if (!object) {
		return null;
	}
	const chart = object.chart;
	const ro = store.readOnly;
	const state = chartElementState(chart);
	const current = chartTypeTargetOf(chart);
	const hasAxes = chart.axes.length > 0;

	const choices = <K extends keyof ChartElementOptions>(element: K, list: Choice<K>): MenuItem[] =>
		list.map(([option, key]) => ({
			id: `${element}-${option}`,
			label: t(key),
			checked: state[element] === option,
			onSelect: () => setSelectedChartElement(store, element, option),
		}));
	const toggle = (element: 'gridlines' | 'axisTitles', key: TranslationKey): MenuItem => ({
		id: element,
		label: t(key),
		checked: state[element] === 'on',
		onSelect: () => setSelectedChartElement(store, element, state[element] === 'on' ? 'off' : 'on'),
	});
	const elementItems: MenuItem[] = [
		{
			id: 'title',
			label: t('chart.element.chartTitle'),
			items: choices('chartTitle', TITLE_CHOICES),
		},
		{ id: 'legend', label: t('chart.element.legend'), items: choices('legend', LEGEND_CHOICES) },
		{
			id: 'labels',
			label: t('chart.element.dataLabels'),
			items: choices('dataLabels', LABEL_CHOICES),
		},
		{
			id: 'gridlines',
			label: t('chart.element.gridlines'),
			disabled: !hasAxes,
			items: [toggle('gridlines', 'chart.option.majorGridlines')],
		},
		{
			id: 'axisTitles',
			label: t('chart.element.axisTitles'),
			disabled: !hasAxes,
			items: [toggle('axisTitles', 'chart.option.axisTitles')],
		},
	];
	const layoutItems: MenuItem[] = Array.from({ length: CHART_QUICK_LAYOUT_COUNT }, (_, i) => ({
		id: `layout-${i + 1}`,
		label: t('chart.layoutN', { n: i + 1 }),
		onSelect: () => applySelectedChartLayout(store, i + 1),
	}));

	return (
		<>
			<RibbonGroup label={t('chart.group.layouts')}>
				<MenuButton
					icon={<SquarePlus size={15} />}
					label={t('chart.addElement')}
					items={elementItems}
					disabled={ro}
				/>
				<MenuButton
					icon={<LayoutDashboard size={15} />}
					label={t('chart.quickLayout')}
					items={layoutItems}
					disabled={ro}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('chart.group.styles')}>
				<MenuButton
					icon={<Palette size={15} />}
					label={t('chart.changeColors')}
					disabled={ro}
					renderPanel={(close) => (
						<ChartColorsGallery
							onPick={(palette) => {
								applySelectedChartPalette(store, palette);
								close();
							}}
						/>
					)}
				/>
				<MenuButton
					icon={<Paintbrush size={15} />}
					label={t('chart.styles')}
					disabled={ro}
					renderPanel={(close) => (
						<ChartStylesGallery
							chart={chart}
							onPick={(style) => {
								applySelectedChartStyle(store, style);
								close();
							}}
						/>
					)}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('chart.group.data')}>
				<RibbonButton
					icon={<ArrowLeftRight size={15} />}
					label={t('chart.switchRowColumn')}
					showLabel
					disabled={ro}
					onClick={() => {
						if (!switchSelectedChartRowColumn(store)) {
							store.setUi({ message: t('chart.switchUnavailable') });
						}
					}}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('chart.group.type')}>
				<MenuButton
					icon={current ? <ChartTypeGlyph type={current} size={20} /> : <ChartColumn size={15} />}
					label={t('chart.changeType')}
					disabled={ro}
					renderPanel={(close) => (
						<ChartTypeGallery
							current={current}
							onPick={(type) => {
								changeChartType(store, type);
								close();
							}}
						/>
					)}
				/>
			</RibbonGroup>
		</>
	);
}
