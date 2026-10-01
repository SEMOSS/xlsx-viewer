import type { CfRule, DifferentialFormat, StyleSheet, Theme } from 'xlsx-viewer-core';
import { dxfBackgroundColor, resolveColorCss } from 'xlsx-viewer-core';
import type { CSSProperties } from 'react';

import { IconSetIcon } from '../../cf/IconSetIcon';

function dxfOf(rule: CfRule, styles: StyleSheet): DifferentialFormat | undefined {
	return rule.dxf ?? (rule.dxfId !== undefined ? styles.dxfs[rule.dxfId] : undefined);
}

function dxfCss(dxf: DifferentialFormat | undefined, theme: Theme): CSSProperties {
	const ctx = { theme };
	const font = dxf?.font;
	const edge = dxf?.border?.top ?? dxf?.border?.left;
	const underline = font?.underline && font.underline !== 'none';
	return {
		backgroundColor: resolveColorCss(dxfBackgroundColor(dxf?.fill), ctx, '#ffffff'),
		color: resolveColorCss(font?.color, ctx, '#000000'),
		fontWeight: font?.bold ? 700 : 400,
		fontStyle: font?.italic ? 'italic' : 'normal',
		textDecorationLine:
			[underline ? 'underline' : '', font?.strike ? 'line-through' : '']
				.filter(Boolean)
				.join(' ') || 'none',
		border:
			edge && edge.style !== 'none'
				? `1px solid ${resolveColorCss(edge.color, ctx, '#000000')}`
				: '1px solid transparent',
	};
}

/** The Rules Manager's "Format" column: what the rule paints. */
export function RuleFormatSample({
	rule,
	styles,
	theme,
}: {
	rule: CfRule;
	styles: StyleSheet;
	theme: Theme;
}) {
	const ctx = { theme };
	if (rule.type === 'colorScale' && rule.colorScale) {
		const stops = rule.colorScale.colors.map((c) => resolveColorCss(c, ctx, '#ffffff')).join(', ');
		return (
			<span
				className='block h-4 w-[96px] border border-black/10'
				style={{ backgroundImage: `linear-gradient(90deg, ${stops})` }}
			/>
		);
	}
	if (rule.type === 'dataBar' && rule.dataBar) {
		const color = resolveColorCss(rule.dataBar.color, ctx, '#638EC6');
		const fill =
			rule.dataBar.gradient === false ? color : `linear-gradient(90deg, ${color}, #ffffff)`;
		return (
			<span className='block h-4 w-[96px] border border-black/10 bg-white p-[2px]'>
				<span
					className='block h-full w-[70%]'
					style={{
						background: fill,
						outline: rule.dataBar.border ? `1px solid ${color}` : undefined,
					}}
				/>
			</span>
		);
	}
	if (rule.type === 'iconSet' && rule.iconSet) {
		const name = rule.iconSet.iconSet;
		const count = rule.iconSet.customIcons?.length || Number.parseInt(name, 10) || 3;
		const order = Array.from({ length: count }, (_, i) =>
			rule.iconSet?.reverse ? i : count - 1 - i,
		);
		return (
			<span className='flex items-center gap-[2px]'>
				{order.map((i) => (
					<IconSetIcon key={i} iconSet={name} index={i} size={14} />
				))}
			</span>
		);
	}
	return (
		<span
			className='inline-block px-1.5 text-[12px] leading-[18px]'
			style={dxfCss(dxfOf(rule, styles), theme)}
		>
			AaBbCcYyZz
		</span>
	);
}
