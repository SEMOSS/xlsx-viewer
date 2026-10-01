import type { TableStyleInfo } from 'xlsx-viewer-core';
import { BUILTIN_TABLE_STYLES, getTableStyle } from 'xlsx-viewer-core';
import { useMemo } from 'react';

import { useT, type Translate, type TranslationKey } from '../../../i18n';
import { buildTableThumbnail, type TableThumbnail } from '../../../engine/render/table-thumbnail';
import { useViewerStore } from '../../state/use-store';
import { TableThumbnailView } from './TableThumbnailView';

/** `TableStyleMedium2` -> `Table Style Medium 2` (other names, such as custom styles, as is). */
export function tableStyleLabel(name: string, t: Translate): string {
	const m = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
	if (!m) {
		return name;
	}
	const key: TranslationKey =
		m[1] === 'Light'
			? 'table.styleLight'
			: m[1] === 'Medium'
				? 'table.styleMedium'
				: 'table.styleDark';
	return t(key, { n: m[2] });
}

interface Entry {
	name: string;
	label: string;
	thumb?: TableThumbnail;
}

interface TableStyleGalleryProps {
	/** Style of the table being edited (highlighted). */
	current?: string;
	/** Table Style Options the previews follow. */
	options?: Partial<TableStyleInfo>;
	/** Offer "None" (clear the style) first among the light styles. */
	allowNone?: boolean;
	onPick: (styleName: string) => void;
}

/** Format as Table / Table Styles gallery: custom, light, medium and dark table styles. */
export function TableStyleGallery({
	current,
	options = {},
	allowNone,
	onPick,
}: TableStyleGalleryProps) {
	const t = useT();
	const store = useViewerStore();
	const styles = store.workbook.styles;
	const theme = store.workbook.theme;
	const optionKey = JSON.stringify(options);

	const groups = useMemo(() => {
		const ctx = { theme, indexedColors: styles.indexedColors };
		const opts = JSON.parse(optionKey) as Partial<TableStyleInfo>;
		const entry = (name: string): Entry => {
			const def = getTableStyle(name, styles);
			return {
				name,
				label: tableStyleLabel(name, t),
				thumb: def ? buildTableThumbnail(def, opts, ctx) : undefined,
			};
		};
		const custom = (styles.customTableStyles ?? [])
			.filter((s) => s.table !== false)
			.map((s) => entry(s.name));
		const builtin = (kind: string) =>
			BUILTIN_TABLE_STYLES.filter((n) => n.startsWith(`TableStyle${kind}`)).map(entry);
		const light = builtin('Light');
		if (allowNone) {
			light.unshift({ name: '', label: t('table.none') });
		}
		return [
			{ key: 'custom', label: t('table.custom'), entries: custom },
			{ key: 'light', label: t('table.light'), entries: light },
			{ key: 'medium', label: t('table.medium'), entries: builtin('Medium') },
			{ key: 'dark', label: t('table.dark'), entries: builtin('Dark') },
		].filter((g) => g.entries.length > 0);
	}, [styles, theme, optionKey, allowNone, t]);

	return (
		<div
			className='xv-table-gallery max-h-[420px] w-[468px] max-w-[calc(100vw-16px)] overflow-y-auto p-2'
			role='listbox'
			aria-label={t('cmd.formatAsTable')}
		>
			{groups.map((g) => (
				<div key={g.key} role='group' aria-label={g.label} className='mb-1.5'>
					<div className='px-1 pb-1 text-[11.5px] font-semibold text-muted-foreground'>
						{g.label}
					</div>
					<div className='grid grid-cols-7 gap-1'>
						{g.entries.map((e) => (
							<button
								key={e.name || 'none'}
								type='button'
								role='option'
								aria-selected={e.name === (current ?? '')}
								aria-label={e.label}
								title={e.label}
								className={`flex h-[50px] items-center justify-center rounded-sm border ${
									e.name === (current ?? '')
										? 'border-primary bg-accent'
										: 'border-transparent hover:border-primary/60 hover:bg-accent/60'
								}`}
								onClick={() => onPick(e.name)}
							>
								{e.thumb ? <TableThumbnailView thumb={e.thumb} /> : <NoneThumbnail />}
							</button>
						))}
					</div>
				</div>
			))}
		</div>
	);
}

function NoneThumbnail() {
	return (
		<svg width={57} height={42} viewBox='0 0 57 42' aria-hidden>
			<rect x={0.5} y={0.5} width={56} height={41} fill='#ffffff' stroke='#c8c8c8' />
			{[0, 1, 2, 3, 4].map((r) => (
				<line
					key={r}
					x1={3}
					x2={54}
					y1={4.5 + r * 8}
					y2={4.5 + r * 8}
					stroke='#9a9a9a'
					strokeDasharray='8 3'
				/>
			))}
		</svg>
	);
}
