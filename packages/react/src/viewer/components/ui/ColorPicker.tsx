import type { ColorSpec, Theme } from 'xlsx-viewer-core';
import { resolveColorHex, themeColorByIndex } from 'xlsx-viewer-core';
import { useId } from 'react';

import { formatPercent, useLocale, useT } from '../../../i18n';

/** Excel's standard colors row. */
const STANDARD = [
	'C00000',
	'FF0000',
	'FFC000',
	'FFFF00',
	'92D050',
	'00B050',
	'00B0F0',
	'0070C0',
	'002060',
	'7030A0',
];

/** Theme columns in Excel's palette order: Background 1, Text 1, Background 2, Text 2, Accent 1-6. */
const THEME_COLUMNS = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];

/** Tint rows shown under each theme color (Excel's 5 variations). */
function tintsFor(index: number, theme: Theme): number[] {
	const hex = themeColorByIndex(theme, index) ?? 'FFFFFF';
	const r = parseInt(hex.slice(0, 2), 16);
	const g = parseInt(hex.slice(2, 4), 16);
	const b = parseInt(hex.slice(4, 6), 16);
	const lum = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
	if (lum > 0.95) {
		return [-0.05, -0.15, -0.25, -0.35, -0.5];
	}
	if (lum < 0.05) {
		return [0.5, 0.35, 0.25, 0.15, 0.05];
	}
	return [0.8, 0.6, 0.4, -0.25, -0.5];
}

interface ColorPickerProps {
	theme: Theme;
	onPick: (color: ColorSpec | null) => void;
	/** Label of the "none" option: Automatic for fonts, No Fill for fills. */
	noneLabel: string;
	close: () => void;
}

/** Office color palette: theme colors with tints, standard colors, none, custom. */
export function ColorPicker({ theme, onPick, noneLabel, close }: ColorPickerProps) {
	const t = useT();
	const locale = useLocale();
	const id = useId();
	const pick = (c: ColorSpec | null) => {
		onPick(c);
		close();
	};
	const swatch = (spec: ColorSpec, title: string, key: string) => {
		const hex = resolveColorHex(spec, { theme }) ?? 'FFFFFF';
		return (
			<button
				key={key}
				type='button'
				title={title}
				aria-label={title}
				className='h-[15px] w-[15px] border border-black/10 hover:z-10 hover:scale-110 hover:outline hover:outline-1 hover:outline-primary'
				style={{ backgroundColor: `#${hex}` }}
				onClick={() => pick(spec)}
			/>
		);
	};
	return (
		<div className='flex w-[208px] flex-col gap-2 p-2 text-[12px]'>
			<button
				type='button'
				className='flex items-center gap-2 rounded-sm px-1 py-1 text-left hover:bg-accent'
				onClick={() => pick(null)}
			>
				<span className='h-4 w-4 border border-border bg-[linear-gradient(135deg,transparent_45%,#e00_45%,#e00_55%,transparent_55%)]' />
				{noneLabel}
			</button>
			<div className='text-[11px] font-semibold text-muted-foreground'>
				{t('color.themeColors')}
			</div>
			<div className='grid grid-cols-10 gap-[3px]'>
				{THEME_COLUMNS.map((i) => swatch({ theme: i }, t('ui.themeColor', { index: i }), `t${i}`))}
			</div>
			<div className='grid grid-cols-10 gap-x-[3px]'>
				{[0, 1, 2, 3, 4].map((row) =>
					THEME_COLUMNS.map((i) => {
						const tint = tintsFor(i, theme)[row];
						const percent = formatPercent(Math.round(Math.abs(tint) * 100), locale);
						const label = t(tint > 0 ? 'ui.themeColorLighter' : 'ui.themeColorDarker', {
							index: i,
							percent,
						});
						return swatch({ theme: i, tint }, label, `t${i}-${row}`);
					}),
				)}
			</div>
			<div className='text-[11px] font-semibold text-muted-foreground'>
				{t('color.standardColors')}
			</div>
			<div className='grid grid-cols-10 gap-[3px]'>
				{STANDARD.map((hex) => swatch({ rgb: `FF${hex}` }, `#${hex}`, hex))}
			</div>
			<label
				htmlFor={id}
				className='flex cursor-pointer items-center gap-2 rounded-sm px-1 py-1 hover:bg-accent'
			>
				<input
					id={id}
					type='color'
					className='h-4 w-4 cursor-pointer border-0 bg-transparent p-0'
					onChange={(e) => pick({ rgb: `FF${e.target.value.slice(1).toUpperCase()}` })}
				/>
				{t('color.more')}
			</label>
		</div>
	);
}
