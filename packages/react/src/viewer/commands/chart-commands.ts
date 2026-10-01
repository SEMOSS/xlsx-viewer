/**
 * Chart Design commands: each edit replaces the selected chart's model with
 * the result of a pure core helper, as one undoable step.
 */
import type {
	ChartColorPalette,
	ChartElementKind,
	ChartElementOptions,
	ChartModel,
	ChartObject,
	ChartTypeTarget,
} from 'xlsx-viewer-core';
import {
	applyChartStyle,
	applyColorPalette,
	applyQuickLayout,
	convertChartType,
	setChartElement,
	switchRowColumn,
} from 'xlsx-viewer-core';

import type { ViewerStore } from '../state/viewer-store';

/** The selected chart object on the active sheet, if a chart is selected. */
export function selectedChart(store: ViewerStore): ChartObject | null {
	const id = store.state.selectedObject;
	if (!id) {
		return null;
	}
	const object = store.activeSheet?.drawings.find((d) => d.id === id);
	return object?.kind === 'chart' ? object : null;
}

/**
 * Replace the selected chart's model with `edit(model)` as one undoable step.
 * Returns false when nothing is selected, the edit does not apply (returns
 * `null` or the same model) or the workbook is read-only.
 */
export function editSelectedChart(
	store: ViewerStore,
	label: string,
	edit: (chart: ChartModel) => ChartModel | null,
): boolean {
	const object = selectedChart(store);
	if (!object) {
		return false;
	}
	const next = edit(object.chart);
	if (!next || next === object.chart) {
		return false;
	}
	const sheetIndex = store.state.activeSheet;
	return store.transact(label, (tx) => {
		tx.touchSheet(sheetIndex, false, 'objects');
		const sheet = store.workbook.sheets[sheetIndex];
		sheet.drawings = sheet.drawings.map((d) =>
			d.id === object.id && d.kind === 'chart' ? { ...d, chart: { ...next, dirty: true } } : d,
		);
		sheet.drawingsDirty = true;
		sheet.dirty = true;
	});
}

export function changeChartType(store: ViewerStore, target: ChartTypeTarget): boolean {
	return editSelectedChart(store, store.t('chart.changeType'), (chart) =>
		convertChartType(chart, target),
	);
}

export function setSelectedChartElement<K extends ChartElementKind>(
	store: ViewerStore,
	element: K,
	option: ChartElementOptions[K],
): boolean {
	return editSelectedChart(store, store.t('chart.addElement'), (chart) =>
		setChartElement(chart, element, option),
	);
}

export function applySelectedChartLayout(store: ViewerStore, layout: number): boolean {
	return editSelectedChart(store, store.t('chart.quickLayout'), (chart) =>
		applyQuickLayout(chart, layout),
	);
}

export function applySelectedChartPalette(store: ViewerStore, palette: ChartColorPalette): boolean {
	return editSelectedChart(store, store.t('chart.changeColors'), (chart) =>
		applyColorPalette(chart, palette),
	);
}

export function applySelectedChartStyle(store: ViewerStore, style: number): boolean {
	return editSelectedChart(store, store.t('undo.chartStyle'), (chart) =>
		applyChartStyle(chart, style),
	);
}

/** Switch Row/Column; false when the chart's data cannot be switched. */
export function switchSelectedChartRowColumn(store: ViewerStore): boolean {
	const sheetIndex = store.state.activeSheet;
	return editSelectedChart(store, store.t('chart.switchRowColumn'), (chart) =>
		switchRowColumn(chart, store.workbook, sheetIndex),
	);
}
