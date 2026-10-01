import type { IconSetName } from 'xlsx-viewer-core';
import {
	COLOR_SCALE_PRESETS,
	DATA_BAR_PRESETS,
	ICON_SET_GROUPS,
	ICON_SETS,
} from 'xlsx-viewer-core';
import { useId, type ReactNode } from 'react';

import { useT } from '../../../../i18n';
import { IconSetIcon } from '../../cf/IconSetIcon';
import {
	COLOR_SCALE_KEYS,
	DATA_BAR_KEYS,
	ICON_SET_GROUP_KEYS,
	ICON_SET_KEYS,
	presetLabel,
} from '../../cf/cf-labels';

function css(argb: string): string {
	return `#${argb.length === 8 ? argb.slice(2) : argb}`;
}

function mix(a: string, b: string): string {
	const pa = parseInt(a.slice(-6), 16);
	const pb = parseInt(b.slice(-6), 16);
	const ch = (shift: number) => Math.round((((pa >> shift) & 255) + ((pb >> shift) & 255)) / 2);
	return `#${[16, 8, 0].map((s) => ch(s).toString(16).padStart(2, '0')).join('')}`;
}

function GallerySection({ title, children }: { title: string; children: ReactNode }) {
	return (
		<div className='flex flex-col gap-1.5'>
			<div className='bg-muted/70 px-2 py-1 text-[12px] font-semibold text-muted-foreground'>
				{title}
			</div>
			<div className='flex flex-wrap gap-1.5 px-2 pb-1'>{children}</div>
		</div>
	);
}

function Tile({
	label,
	onClick,
	children,
	wide,
}: {
	label: string;
	onClick: () => void;
	children: ReactNode;
	wide?: boolean;
}) {
	return (
		<button
			type='button'
			role='menuitem'
			title={label}
			aria-label={label}
			className={`flex h-[42px] ${wide ? 'w-[112px]' : 'w-[42px]'} items-center justify-center rounded-sm border border-transparent outline-none hover:border-primary hover:bg-accent focus-visible:border-primary focus-visible:bg-accent`}
			onClick={onClick}
		>
			{children}
		</button>
	);
}

/** Four bars of decreasing length, like Excel's data bar gallery icons. */
function DataBarPreview({ color, gradient }: { color: string; gradient: boolean }) {
	const id = `xv-dbg-${useId().replace(/[^A-Za-z0-9_-]/g, '')}`;
	const fill = gradient ? `url(#${id})` : css(color);
	return (
		<svg width={32} height={32} viewBox='0 0 32 32' aria-hidden='true'>
			{gradient && (
				<defs>
					<linearGradient id={id} x1={0} x2={1} y1={0} y2={0}>
						<stop offset='0' stopColor={css(color)} />
						<stop offset='1' stopColor='#ffffff' />
					</linearGradient>
				</defs>
			)}
			<rect x={0.5} y={0.5} width={31} height={31} fill='#ffffff' stroke='#d0d0d0' />
			{[28, 21, 14, 7].map((w, i) => (
				<rect
					key={w}
					x={2}
					y={3 + i * 7}
					width={w}
					height={5}
					fill={fill}
					stroke={gradient ? css(color) : undefined}
					strokeWidth={0.75}
				/>
			))}
		</svg>
	);
}

/** Data Bars: gradient and solid fills in Excel's six colors. */
export function DataBarGallery({ onPick }: { onPick: (color: string, gradient: boolean) => void }) {
	const t = useT();
	return (
		<div role='menu' aria-label={t('cf.dataBars')} className='flex w-[176px] flex-col gap-2 py-1.5'>
			{([true, false] as const).map((gradient) => (
				<GallerySection
					key={String(gradient)}
					title={t(gradient ? 'cf.gradientFill' : 'cf.solidFill')}
				>
					{DATA_BAR_PRESETS.map((p) => (
						<Tile
							key={p.name}
							label={t(gradient ? 'cf.dataBarGradient' : 'cf.dataBarSolid', {
								name: presetLabel(DATA_BAR_KEYS, p.name, t),
							})}
							onClick={() => onPick(p.color, gradient)}
						>
							<DataBarPreview color={p.color} gradient={gradient} />
						</Tile>
					))}
				</GallerySection>
			))}
		</div>
	);
}

/** A 3x3 grid colored from the highest values (top) to the lowest (bottom). */
function ColorScalePreview({ colors }: { colors: readonly string[] }) {
	const high = css(colors[colors.length - 1]);
	const low = css(colors[0]);
	const mid = colors.length === 3 ? css(colors[1]) : mix(colors[0], colors[1]);
	return (
		<svg width={32} height={32} viewBox='0 0 32 32' aria-hidden='true'>
			{[high, mid, low].map((c, row) =>
				[0, 1, 2].map((col) => (
					<rect
						key={`${row}-${col}`}
						x={1 + col * 10}
						y={1 + row * 10}
						width={10}
						height={10}
						fill={c}
						stroke='#ffffff'
						strokeWidth={0.75}
					/>
				)),
			)}
		</svg>
	);
}

/** Color Scales: Excel's twelve presets. */
export function ColorScaleGallery({ onPick }: { onPick: (presetIndex: number) => void }) {
	const t = useT();
	return (
		<div
			role='menu'
			aria-label={t('cf.colorScales')}
			className='flex w-[200px] flex-wrap gap-1.5 p-2'
		>
			{COLOR_SCALE_PRESETS.map((p, i) => (
				<Tile
					key={p.name}
					label={presetLabel(COLOR_SCALE_KEYS, p.name, t)}
					onClick={() => onPick(i)}
				>
					<ColorScalePreview colors={p.colors} />
				</Tile>
			))}
		</div>
	);
}

/** Icon Sets grouped like Excel (Directional, Shapes, Indicators, Ratings). */
export function IconSetGallery({ onPick }: { onPick: (name: IconSetName) => void }) {
	const t = useT();
	return (
		<div role='menu' aria-label={t('cf.iconSets')} className='flex w-[262px] flex-col gap-2 py-1.5'>
			{ICON_SET_GROUPS.map((group) => (
				<GallerySection key={group.label} title={presetLabel(ICON_SET_GROUP_KEYS, group.label, t)}>
					{group.sets.map((name) => {
						const set = ICON_SETS[name];
						// Highest value first, as the gallery shows them.
						const order = set.icons.map((_, i) => set.count - 1 - i);
						return (
							<Tile key={name} label={t(ICON_SET_KEYS[name])} onClick={() => onPick(name)} wide>
								<span className='flex items-center gap-[3px]'>
									{order.map((index) => (
										<IconSetIcon key={index} iconSet={name} index={index} size={16} />
									))}
								</span>
							</Tile>
						);
					})}
				</GallerySection>
			))}
		</div>
	);
}
