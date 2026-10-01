import { useId, useMemo, useState } from 'react';
import type { DocumentTextProperty } from 'xlsx-viewer-core';
import { setDocumentProperties, workbookStatistics } from 'xlsx-viewer-core';

import { formatNumber, useLocale, useT, type Translate, type TranslationKey } from '../../../i18n';
import { useViewerStore } from '../../state/use-store';
import { FIELD, FormDialog, TabPanel, TabStrip, type TabDef } from './dialog-ui';
import type { DialogProps } from './types';

type TabId = 'summary' | 'statistics';

const TAB_KEYS: ReadonlyArray<{ id: TabId; labelKey: TranslationKey }> = [
	{ id: 'summary', labelKey: 'properties.summary' },
	{ id: 'statistics', labelKey: 'properties.statistics' },
];

type SummaryField = Exclude<DocumentTextProperty, 'lastModifiedBy'>;

/** The Summary tab, in Excel's order (Author is `creator`, Comments is `description`). */
const SUMMARY_FIELDS: ReadonlyArray<{ key: SummaryField; labelKey: TranslationKey }> = [
	{ key: 'title', labelKey: 'properties.title' },
	{ key: 'subject', labelKey: 'properties.subject' },
	{ key: 'creator', labelKey: 'properties.author' },
	{ key: 'manager', labelKey: 'properties.manager' },
	{ key: 'company', labelKey: 'properties.company' },
	{ key: 'category', labelKey: 'properties.category' },
	{ key: 'keywords', labelKey: 'properties.keywords' },
	{ key: 'description', labelKey: 'properties.comments' },
];

function dateText(iso: string | undefined, t: Translate, locale: string): string {
	if (!iso) {
		return t('properties.notKnown');
	}
	const date = new Date(iso);
	return Number.isNaN(date.getTime()) ? t('properties.notKnown') : date.toLocaleString(locale);
}

/**
 * File > Info > Properties > Advanced Properties: the document's summary
 * properties and statistics. Changes are not undoable, like in Excel.
 */
export function PropertiesDialog({ onClose }: DialogProps) {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const idBase = useId();
	const props = store.workbook.properties;
	const [tab, setTab] = useState<TabId>('summary');
	const [values, setValues] = useState<Record<SummaryField, string>>(() => {
		const out = {} as Record<SummaryField, string>;
		for (const { key } of SUMMARY_FIELDS) {
			out[key] = props[key] ?? '';
		}
		return out;
	});
	const stats = useMemo(() => workbookStatistics(store.workbook), [store]);
	const readOnly = store.readOnly;

	const ok = () => {
		if (!readOnly && setDocumentProperties(store.workbook, values)) {
			store.touch();
		}
		onClose();
	};

	const tabs: TabDef<TabId>[] = TAB_KEYS.map(({ id, labelKey }) => ({ id, label: t(labelKey) }));
	const statistics: Array<[string, string]> = [
		[t('properties.created'), dateText(props.created, t, locale)],
		[t('properties.modified'), dateText(props.modified, t, locale)],
		[t('properties.lastModifiedBy'), props.lastModifiedBy || t('properties.notKnown')],
		[t('properties.sheets'), formatNumber(stats.sheets, locale)],
		[t('properties.cellsWithValues'), formatNumber(stats.cells, locale)],
		[t('properties.formulas'), formatNumber(stats.formulas, locale)],
	];

	return (
		<FormDialog
			title={t('file.properties')}
			onClose={onClose}
			onOk={ok}
			width={420}
			initialFocusSelector='input'
		>
			<TabStrip
				tabs={tabs}
				value={tab}
				onChange={setTab}
				label={t('file.properties')}
				idBase={idBase}
			/>
			<TabPanel idBase={idBase} id={tab} className='min-h-[300px]'>
				{tab === 'summary' ? (
					<div className='grid grid-cols-[88px_minmax(0,1fr)] items-center gap-x-3 gap-y-1.5'>
						{SUMMARY_FIELDS.map(({ key, labelKey }) => {
							const id = `${idBase}-${key}`;
							const common = {
								id,
								className: key === 'description' ? `${FIELD} h-[80px] resize-none` : FIELD,
								value: values[key],
								readOnly,
								onChange: (e: { target: { value: string } }) =>
									setValues((v) => ({ ...v, [key]: e.target.value })),
							};
							return (
								<div key={key} className='contents'>
									<label
										htmlFor={id}
										className={key === 'description' ? 'self-start pt-1' : undefined}
									>
										{t(labelKey)}
									</label>
									{key === 'description' ? <textarea {...common} /> : <input {...common} />}
								</div>
							);
						})}
					</div>
				) : (
					<dl className='grid grid-cols-[130px_minmax(0,1fr)] gap-x-3 gap-y-2'>
						{statistics.map(([label, value]) => (
							<div key={label} className='contents'>
								<dt className='text-muted-foreground'>{label}</dt>
								<dd className='break-words'>{value}</dd>
							</div>
						))}
					</dl>
				)}
			</TabPanel>
		</FormDialog>
	);
}
