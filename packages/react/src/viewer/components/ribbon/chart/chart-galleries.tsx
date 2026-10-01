/** Galleries of the Chart Design tab: chart types, color palettes and quick styles. */
import type { ChartColorPalette, ChartModel, ChartTypeTarget } from 'xlsx-viewer-core';
import {
	applyChartStyle,
	CHART_COLOR_PALETTES,
	CHART_QUICK_STYLE_COUNT,
	chartPaletteColor,
	layoutChart,
	resolveChartData,
	resolveDrawingColorCss,
} from 'xlsx-viewer-core';
import { useMemo } from 'react';

import { useT, type TranslationKey } from '../../../../i18n';
import { useViewerStore } from '../../../state/use-store';
import { ChartSvg } from '../../charts/ChartSvg';
import { ChartTypeGlyph } from './chart-glyphs';

const FAMILIES: Array<{ key: TranslationKey; types: ChartTypeTarget[] }> = [
	{ key: 'chart.family.column', types: ['columnClustered', 'columnStacked', 'columnPercent'] },
	{ key: 'chart.family.bar', types: ['barClustered', 'barStacked', 'barPercent'] },
	{ key: 'chart.family.line', types: ['line', 'lineMarkers'] },
	{ key: 'chart.family.area', types: ['area', 'areaStacked'] },
	{ key: 'chart.family.pie', types: ['pie', 'doughnut'] },
	{ key: 'chart.family.scatter', types: ['scatter', 'scatterLines'] },
	{ key: 'chart.family.radar', types: ['radar'] },
];

const optionClass = (selected: boolean) =>
	`flex items-center justify-center rounded-sm border ${
		selected
			? 'border-primary bg-accent'
			: 'border-transparent hover:border-primary/60 hover:bg-accent/60'
	}`;

/** Change Chart Type: one row of type pictures per chart family. */
export function ChartTypeGallery({
	current,
	onPick,
}: {
	current?: ChartTypeTarget;
	onPick: (type: ChartTypeTarget) => void;
}) {
	const t = useT();
	return (
		<div className='w-[300px] p-2' role='listbox' aria-label={t('chart.changeType')}>
			{FAMILIES.map((family) => (
				<div
					key={family.key}
					role='group'
					aria-label={t(family.key)}
					className='mb-1 flex items-center gap-1'
				>
					<div className='w-[92px] shrink-0 px-1 text-[11.5px] text-muted-foreground'>
						{t(family.key)}
					</div>
					{family.types.map((type) => {
						const label = t(`chart.type.${type}`);
						return (
							<button
								key={type}
								type='button'
								role='option'
								aria-selected={type === current}
								aria-label={label}
								title={label}
								className={`${optionClass(type === current)} h-[34px] w-[46px]`}
								onClick={() => onPick(type)}
							>
								<ChartTypeGlyph type={type} />
							</button>
						);
					})}
				</div>
			))}
		</div>
	);
}

/** Change Colors: the colorful and monochromatic palettes as rows of swatches. */
export function ChartColorsGallery({ onPick }: { onPick: (palette: ChartColorPalette) => void }) {
	const t = useT();
	const theme = useViewerStore().workbook.theme;
	const sections: Array<{ key: TranslationKey; palettes: ChartColorPalette[] }> = [
		{
			key: 'chart.colorful',
			palettes: CHART_COLOR_PALETTES.filter((p) => p.startsWith('colorful')),
		},
		{
			key: 'chart.monochromatic',
			palettes: CHART_COLOR_PALETTES.filter((p) => p.startsWith('monochromatic')),
		},
	];
	return (
		<div className='w-[220px] p-2' role='listbox' aria-label={t('chart.changeColors')}>
			{sections.map((section) => (
				<div key={section.key} role='group' aria-label={t(section.key)} className='mb-1.5'>
					<div className='px-1 pb-1 text-[11.5px] font-semibold text-muted-foreground'>
						{t(section.key)}
					</div>
					{section.palettes.map((palette) => {
						const n = Number(palette.replace(/\D/g, ''));
						const label = t(
							palette.startsWith('colorful') ? 'chart.paletteColorful' : 'chart.paletteMono',
							{ n },
						);
						return (
							<button
								key={palette}
								type='button'
								role='option'
								aria-selected={false}
								aria-label={label}
								title={label}
								className={`${optionClass(false)} w-full justify-start gap-0.5 px-1 py-1`}
								onClick={() => onPick(palette)}
							>
								{Array.from({ length: 6 }, (_, i) => (
									<span
										key={i}
										className='h-3.5 w-6 rounded-[1px]'
										style={{
											background: resolveDrawingColorCss(chartPaletteColor(palette, i, 6), theme),
										}}
									/>
								))}
							</button>
						);
					})}
				</div>
			))}
		</div>
	);
}

/** Chart Styles: previews of the chart in each quick style. */
export function ChartStylesGallery({
	chart,
	onPick,
}: {
	chart: ChartModel;
	onPick: (style: number) => void;
}) {
	const t = useT();
	const store = useViewerStore();
	const sheetIndex = store.state.activeSheet;
	const theme = store.workbook.theme;
	const previews = useMemo(() => {
		const data = resolveChartData(chart, store.workbook, sheetIndex);
		return Array.from({ length: CHART_QUICK_STYLE_COUNT }, (_, i) =>
			layoutChart(applyChartStyle(chart, i + 1), data, theme, 240, 160),
		);
	}, [chart, store.workbook, sheetIndex, theme]);
	return (
		<div
			className='grid w-[392px] grid-cols-3 gap-1 p-2'
			role='listbox'
			aria-label={t('chart.styles')}
		>
			{previews.map((tree, i) => {
				const label = t('chart.styleN', { n: i + 1 });
				return (
					<button
						key={label}
						type='button'
						role='option'
						aria-selected={false}
						aria-label={label}
						title={label}
						className={`${optionClass(false)} h-[86px] p-0.5`}
						onClick={() => onPick(i + 1)}
					>
						<ChartSvg tree={tree} style={{ width: 120, height: 80, pointerEvents: 'none' }} />
					</button>
				);
			})}
		</div>
	);
}
