import type { BorderPreset, ColorSpec } from 'xlsx-viewer-core';
import { resolveColorCss } from 'xlsx-viewer-core';
import {
	AArrowDown,
	AArrowUp,
	Baseline,
	Bold,
	Grid2x2,
	Italic,
	PaintBucket,
	Square,
	SquareDashed,
	Strikethrough,
	Underline,
} from 'lucide-react';
import { useState } from 'react';

import { useT } from '../../../../i18n';
import {
	applyBorders,
	setFillColor,
	setFontColor,
	setFontName,
	setFontSize,
	stepFontSize,
	toggleFontFlag,
	toggleUnderline,
} from '../../../commands/format-commands';
import { useActiveFormat } from '../../../hooks/useActiveFormat';
import { useViewerStore } from '../../../state/use-store';
import { ColorPicker } from '../../ui/ColorPicker';
import { ComboBox } from '../../ui/ComboBox';
import { ControlCluster, RibbonButton, SplitButton } from '../../ui/buttons';
import type { MenuItem } from '../../ui/Menu';
import { RibbonGroup } from '../RibbonGroup';

export const FONT_NAMES = [
	'Aptos',
	'Aptos Narrow',
	'Arial',
	'Arial Narrow',
	'Calibri',
	'Calibri Light',
	'Cambria',
	'Candara',
	'Century Gothic',
	'Consolas',
	'Courier New',
	'Franklin Gothic Medium',
	'Garamond',
	'Georgia',
	'Helvetica',
	'Segoe UI',
	'Tahoma',
	'Times New Roman',
	'Trebuchet MS',
	'Verdana',
];

export const FONT_SIZES = [
	'8',
	'9',
	'10',
	'10.5',
	'11',
	'12',
	'14',
	'16',
	'18',
	'20',
	'22',
	'24',
	'26',
	'28',
	'36',
	'48',
	'72',
];

/** Swatch bar under a color button icon. */
function ColorIcon({ icon, color }: { icon: React.ReactNode; color: string }) {
	return (
		<span className='relative inline-flex flex-col items-center'>
			{icon}
			<span
				className='mt-[1px] h-[3px] w-[15px] rounded-[1px]'
				style={{ backgroundColor: color }}
			/>
		</span>
	);
}

export function FontGroup() {
	const t = useT();
	const store = useViewerStore();
	const { style } = useActiveFormat();
	const readOnly = store.readOnly;
	const theme = store.workbook.theme;
	const [fill, setFill] = useState<ColorSpec>({ rgb: 'FFFFFF00' });
	const [fontColor, setFontColorState] = useState<ColorSpec>({ rgb: 'FFFF0000' });
	const [border, setBorder] = useState<BorderPreset>('bottom');
	const borderItems: MenuItem[] = (
		[
			['bottom', 'border.bottom'],
			['top', 'border.top'],
			['left', 'border.left'],
			['right', 'border.right'],
			['none', 'border.none'],
			['all', 'border.all'],
			['outside', 'border.outside'],
			['thickOutside', 'border.thickOutside'],
			['bottomDouble', 'border.bottomDouble'],
			['thickBottom', 'border.thickBottom'],
			['topBottom', 'border.topBottom'],
			['topThickBottom', 'border.topThickBottom'],
			['topDoubleBottom', 'border.topDoubleBottom'],
		] as const
	).map(([preset, key]) => ({
		label: t(key),
		icon:
			preset === 'none' ? (
				<SquareDashed size={14} />
			) : preset === 'all' ? (
				<Grid2x2 size={14} />
			) : (
				<Square size={14} />
			),
		onSelect: () => {
			setBorder(preset);
			applyBorders(store, preset);
		},
	}));
	return (
		<RibbonGroup label={t('group.font')}>
			<ComboBox
				label={t('cmd.fontName')}
				value={style.font.name}
				options={FONT_NAMES}
				previewFont
				width={128}
				disabled={readOnly}
				onCommit={(v) => setFontName(store, v)}
			/>
			<ComboBox
				label={t('cmd.fontSize')}
				value={String(style.font.size)}
				options={FONT_SIZES}
				width={58}
				disabled={readOnly}
				onCommit={(v) => setFontSize(store, Number(v))}
			/>
			<ControlCluster>
				<RibbonButton
					icon={<AArrowUp size={15} />}
					label={t('cmd.increaseFontSize')}
					disabled={readOnly}
					onClick={() => stepFontSize(store, 1)}
				/>
				<RibbonButton
					icon={<AArrowDown size={15} />}
					label={t('cmd.decreaseFontSize')}
					disabled={readOnly}
					onClick={() => stepFontSize(store, -1)}
				/>
			</ControlCluster>
			<ControlCluster>
				<RibbonButton
					icon={<Bold size={15} />}
					label={t('cmd.bold')}
					shortcut='Ctrl+B'
					pressed={!!style.font.bold}
					disabled={readOnly}
					onClick={() => toggleFontFlag(store, 'bold')}
				/>
				<RibbonButton
					icon={<Italic size={15} />}
					label={t('cmd.italic')}
					shortcut='Ctrl+I'
					pressed={!!style.font.italic}
					disabled={readOnly}
					onClick={() => toggleFontFlag(store, 'italic')}
				/>
				<SplitButton
					icon={<Underline size={15} />}
					label={t('cmd.underline')}
					shortcut='Ctrl+U'
					pressed={!!style.font.underline && style.font.underline !== 'none'}
					disabled={readOnly}
					onClick={() => toggleUnderline(store)}
					items={[
						{ label: t('cmd.underline'), onSelect: () => toggleUnderline(store, 'single') },
						{ label: t('cmd.doubleUnderline'), onSelect: () => toggleUnderline(store, 'double') },
					]}
				/>
				<RibbonButton
					icon={<Strikethrough size={15} />}
					label={t('cmd.strikethrough')}
					shortcut='Ctrl+5'
					pressed={!!style.font.strike}
					disabled={readOnly}
					onClick={() => toggleFontFlag(store, 'strike')}
				/>
			</ControlCluster>
			<SplitButton
				icon={<Square size={15} />}
				label={t('cmd.borders')}
				disabled={readOnly}
				onClick={() => applyBorders(store, border)}
				items={borderItems}
			/>
			<SplitButton
				icon={
					<ColorIcon
						icon={<PaintBucket size={14} />}
						color={resolveColorCss(fill, { theme }) ?? '#ffff00'}
					/>
				}
				label={t('cmd.fillColor')}
				disabled={readOnly}
				onClick={() => setFillColor(store, fill)}
				renderPanel={(close) => (
					<ColorPicker
						theme={theme}
						noneLabel={t('color.noFill')}
						close={close}
						onPick={(c) => {
							if (c) {
								setFill(c);
							}
							setFillColor(store, c);
						}}
					/>
				)}
			/>
			<SplitButton
				icon={
					<ColorIcon
						icon={<Baseline size={14} />}
						color={resolveColorCss(fontColor, { theme }) ?? '#ff0000'}
					/>
				}
				label={t('cmd.fontColor')}
				disabled={readOnly}
				onClick={() => setFontColor(store, fontColor)}
				renderPanel={(close) => (
					<ColorPicker
						theme={theme}
						noneLabel={t('color.automatic')}
						close={close}
						onPick={(c) => {
							if (c) {
								setFontColorState(c);
							}
							setFontColor(store, c ?? undefined);
						}}
					/>
				)}
			/>
		</RibbonGroup>
	);
}
