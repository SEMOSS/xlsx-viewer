import type { CellStyleGroup } from 'xlsx-viewer-core';
import { BUILTIN_CELL_STYLES, getBuiltinCellStyle } from 'xlsx-viewer-core';
import { useMemo } from 'react';

import { useT, type TranslationKey } from '../../../i18n';
import { cellStylePreview } from '../../../engine/render/cell-style-preview';
import { useViewerStore } from '../../state/use-store';

const GROUPS: Array<{ group: CellStyleGroup; key: TranslationKey }> = [
	{ group: 'goodBadNeutral', key: 'cellStyles.goodBadNeutral' },
	{ group: 'dataModel', key: 'cellStyles.dataModel' },
	{ group: 'titles', key: 'cellStyles.titles' },
	{ group: 'themed', key: 'cellStyles.themed' },
	{ group: 'numberFormat', key: 'cellStyles.numberFormat' },
];

/** Home > Cell Styles gallery: custom styles of the workbook, then Excel's built-in groups. */
export function CellStyleGallery({ onPick }: { onPick: (styleName: string) => void }) {
	const t = useT();
	const store = useViewerStore();
	const styles = store.workbook.styles;
	const theme = store.workbook.theme;

	const sections = useMemo(() => {
		const tile = (name: string) => ({ name, preview: cellStylePreview(styles, theme, name, 18) });
		const custom = styles.cellStyles
			.filter(
				(s) =>
					!s.hidden &&
					(s.builtinId === undefined || s.customBuiltin) &&
					!getBuiltinCellStyle(s.name),
			)
			.map((s) => tile(s.name));
		return [
			...(custom.length ? [{ key: 'custom', label: t('cellStyles.custom'), tiles: custom }] : []),
			...GROUPS.map(({ group, key }) => ({
				key: group,
				label: t(key),
				tiles: BUILTIN_CELL_STYLES.filter((s) => s.group === group).map((s) => tile(s.name)),
			})),
		];
	}, [styles, theme, t]);

	return (
		<div
			className='xv-cell-style-gallery max-h-[440px] w-[640px] max-w-[calc(100vw-16px)] overflow-y-auto p-2'
			role='listbox'
			aria-label={t('cmd.cellStyles')}
		>
			{sections.map((sec) => (
				<div key={sec.key} role='group' aria-label={sec.label} className='mb-2'>
					<div className='px-1 pb-1 text-[11.5px] font-semibold text-muted-foreground'>
						{sec.label}
					</div>
					<div className='grid grid-cols-6 gap-1'>
						{sec.tiles.map(({ name, preview }) => (
							<button
								key={name}
								type='button'
								role='option'
								aria-selected={false}
								title={name}
								className='h-[30px] overflow-hidden rounded-[2px] border border-transparent px-1 text-left whitespace-nowrap outline-offset-1 hover:outline hover:outline-2 hover:outline-primary/70'
								style={{
									fontFamily: preview?.fontFamily,
									fontSize: Math.min(preview?.fontSize ?? 14, 18),
									fontWeight: preview?.fontWeight,
									fontStyle: preview?.fontStyle,
									color: preview?.color,
									backgroundColor: preview?.backgroundColor ?? '#ffffff',
									backgroundImage: preview?.backgroundImage,
									borderTop: preview?.borderTop,
									borderBottom: preview?.borderBottom,
									borderLeft: preview?.borderLeft,
									borderRight: preview?.borderRight,
								}}
								onClick={() => onPick(name)}
							>
								{name}
							</button>
						))}
					</div>
				</div>
			))}
		</div>
	);
}
