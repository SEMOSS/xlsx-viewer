import type { FontSpec, Theme, UnderlineStyle } from 'xlsx-viewer-core';
import { resolveColorCss } from 'xlsx-viewer-core';
import { useEffect, useState, type CSSProperties, type ReactNode } from 'react';

import { fontFamilyStack } from '../../../../engine/render/fonts';
import { useT, type Translate, type TranslationKey } from '../../../../i18n';
import { MIXED, type FormatCellsValues } from '../../../commands/format-cells-commands';
import { FONT_NAMES, FONT_SIZES } from '../../ribbon/home/FontGroup';
import { ColorSelect } from '../ColorSelect';
import { Checkbox, FIELD, GroupBox, ListBox } from '../dialog-ui';

type Patch = Partial<FormatCellsValues>;

interface FontEntry {
	key: string;
	label: string;
	name: string;
	scheme?: 'major' | 'minor';
}

type FontStyleKey = 'regular' | 'italic' | 'bold' | 'boldItalic';

const STYLES: ReadonlyArray<{ key: FontStyleKey; bold: boolean; italic: boolean }> = [
	{ key: 'regular', bold: false, italic: false },
	{ key: 'italic', bold: false, italic: true },
	{ key: 'bold', bold: true, italic: false },
	{ key: 'boldItalic', bold: true, italic: true },
];

const STYLE_LABELS: Record<FontStyleKey, TranslationKey> = {
	regular: 'dialog.formatCells.fontStyleName.regular',
	italic: 'dialog.formatCells.fontStyleName.italic',
	bold: 'dialog.formatCells.fontStyleName.bold',
	boldItalic: 'dialog.formatCells.fontStyleName.boldItalic',
};

const UNDERLINES: ReadonlyArray<{ value: UnderlineStyle; labelKey: TranslationKey }> = [
	{ value: 'none', labelKey: 'dialog.none' },
	{ value: 'single', labelKey: 'dialog.formatCells.underlineType.single' },
	{ value: 'double', labelKey: 'dialog.formatCells.underlineType.double' },
	{ value: 'singleAccounting', labelKey: 'dialog.formatCells.underlineType.singleAccounting' },
	{ value: 'doubleAccounting', labelKey: 'dialog.formatCells.underlineType.doubleAccounting' },
];

function fontEntries(theme: Theme, t: Translate): FontEntry[] {
	const minor = theme.fonts.minorLatin || 'Calibri';
	const major = theme.fonts.majorLatin || 'Calibri Light';
	const names = [...new Set([...FONT_NAMES, minor, major])].sort((a, b) => a.localeCompare(b));
	return [
		{
			key: 'minor',
			label: t('dialog.formatCells.fontBody', { name: minor }),
			name: minor,
			scheme: 'minor',
		},
		{
			key: 'major',
			label: t('dialog.formatCells.fontHeadings', { name: major }),
			name: major,
			scheme: 'major',
		},
		...names.map((name) => ({ key: name, label: name, name })),
	];
}

/** A text box above its list (Excel's Font / Font style / Size columns); partial text is kept. */
function ComboColumn({
	label,
	text,
	onText,
	children,
}: {
	label: string;
	text: string;
	onText: (t: string) => void;
	children: ReactNode;
}) {
	const [local, setLocal] = useState(text);
	useEffect(() => setLocal(text), [text]);
	return (
		<div className='flex min-w-0 flex-col gap-1'>
			<label className='flex flex-col gap-1'>
				<span>{label}</span>
				<input
					className={FIELD}
					value={local}
					onChange={(e) => {
						setLocal(e.target.value);
						onText(e.target.value);
					}}
				/>
			</label>
			{children}
		</div>
	);
}

function previewStyle(d: FormatCellsValues, theme: Theme): CSSProperties {
	const size = d.fontSize === MIXED ? 11 : d.fontSize;
	const underline = d.underline === MIXED ? 'none' : d.underline;
	const decorations = [
		underline !== 'none' ? 'underline' : '',
		d.strike === true ? 'line-through' : '',
	].filter(Boolean);
	return {
		fontFamily: fontFamilyStack(d.fontName === MIXED ? undefined : d.fontName),
		fontSize: Math.min(40, Math.max(6, (size * 96) / 72)),
		fontWeight: d.bold === true ? 700 : 400,
		fontStyle: d.italic === true ? 'italic' : 'normal',
		textDecorationLine: decorations.join(' ') || 'none',
		textDecorationStyle:
			underline === 'double' || underline === 'doubleAccounting' ? 'double' : 'solid',
		color: resolveColorCss(d.fontColor === MIXED ? undefined : d.fontColor, { theme }, '#000000'),
	};
}

/** Format Cells > Font. */
export function FontTab({
	draft,
	onChange,
	theme,
	normalFont,
}: {
	draft: FormatCellsValues;
	onChange: (p: Patch) => void;
	theme: Theme;
	normalFont: FontSpec;
}) {
	const t = useT();
	const entries = fontEntries(theme, t);
	const styles = STYLES.map((s) => ({ ...s, label: t(STYLE_LABELS[s.key]) }));
	const entryKey =
		draft.fontScheme === 'minor' || draft.fontScheme === 'major'
			? draft.fontScheme
			: draft.fontName === MIXED
				? null
				: (entries.find(
						(e) => !e.scheme && e.name.toLowerCase() === String(draft.fontName).toLowerCase(),
					)?.key ?? null);
	const entry = entries.find((e) => e.key === entryKey);
	const fontText = entry ? entry.label : draft.fontName === MIXED ? '' : draft.fontName;
	const styleKey =
		draft.bold === MIXED || draft.italic === MIXED
			? null
			: (styles.find((s) => s.bold === draft.bold && s.italic === draft.italic)?.key ?? null);
	const sizeText = draft.fontSize === MIXED ? '' : String(draft.fontSize);
	const normal = {
		fontName: normalFont.name ?? 'Calibri',
		fontScheme: normalFont.scheme === 'none' ? undefined : normalFont.scheme,
		fontSize: normalFont.size ?? 11,
		bold: Boolean(normalFont.bold),
		italic: Boolean(normalFont.italic),
		underline: normalFont.underline ?? 'none',
		fontColor: normalFont.color,
		strike: Boolean(normalFont.strike),
		vertAlign: 'baseline' as const,
	};
	const isNormal = (Object.keys(normal) as Array<keyof typeof normal>).every(
		(k) => JSON.stringify(draft[k] ?? null) === JSON.stringify(normal[k] ?? null),
	);
	const vertAlign = draft.vertAlign;
	return (
		<div className='flex flex-col gap-3'>
			<div className='grid grid-cols-[1fr_128px_72px] gap-3'>
				<ComboColumn
					label={t('dialog.formatCells.fontLabel')}
					text={fontText}
					onText={(text) => {
						const typed = text.trim().toLowerCase();
						const hit = entries.find((e) => e.label.toLowerCase() === typed);
						onChange(
							hit
								? { fontName: hit.name, fontScheme: hit.scheme }
								: { fontName: text, fontScheme: undefined },
						);
					}}
				>
					<ListBox
						label={t('dialog.formatCells.font')}
						items={entries}
						getKey={(e) => e.key}
						getText={(e) => e.label}
						value={entryKey}
						onChange={(key) => {
							const e = entries.find((x) => x.key === key)!;
							onChange({ fontName: e.name, fontScheme: e.scheme });
						}}
						render={(e) => <span style={{ fontFamily: fontFamilyStack(e.name) }}>{e.label}</span>}
						className='h-[118px]'
					/>
				</ComboColumn>
				<ComboColumn
					label={t('dialog.formatCells.fontStyleLabel')}
					text={styles.find((s) => s.key === styleKey)?.label ?? ''}
					onText={(text) => {
						const typed = text.trim().toLowerCase();
						const hit = styles.find((s) => s.label.toLowerCase() === typed);
						if (hit) {
							onChange({ bold: hit.bold, italic: hit.italic });
						}
					}}
				>
					<ListBox
						label={t('dialog.formatCells.fontStyle')}
						items={styles}
						getKey={(s) => s.key}
						getText={(s) => s.label}
						value={styleKey}
						onChange={(key) => {
							const s = styles.find((x) => x.key === key)!;
							onChange({ bold: s.bold, italic: s.italic });
						}}
						render={(s) => (
							<span
								style={{
									fontWeight: s.bold ? 700 : 400,
									fontStyle: s.italic ? 'italic' : 'normal',
								}}
							>
								{s.label}
							</span>
						)}
						className='h-[118px]'
					/>
				</ComboColumn>
				<ComboColumn
					label={t('dialog.formatCells.sizeLabel')}
					text={sizeText}
					onText={(text) => {
						const n = Number(text);
						if (n >= 1 && n <= 409) {
							onChange({ fontSize: Math.round(n * 2) / 2 });
						}
					}}
				>
					<ListBox
						label={t('dialog.formatCells.size')}
						items={FONT_SIZES}
						getKey={(s) => s}
						value={FONT_SIZES.includes(sizeText) ? sizeText : null}
						onChange={(s) => onChange({ fontSize: Number(s) })}
						render={(s) => s}
						className='h-[118px]'
					/>
				</ComboColumn>
			</div>
			<div className='grid grid-cols-[1fr_1fr] items-end gap-3'>
				<label className='flex flex-col gap-1'>
					<span>{t('dialog.formatCells.underline')}</span>
					<select
						className={FIELD}
						value={draft.underline === MIXED ? '' : draft.underline}
						onChange={(e) => onChange({ underline: e.target.value as UnderlineStyle })}
					>
						{draft.underline === MIXED && <option value='' />}
						{UNDERLINES.map((u) => (
							<option key={u.value} value={u.value}>
								{t(u.labelKey)}
							</option>
						))}
					</select>
				</label>
				<div className='flex items-end justify-between gap-2'>
					<div className='flex flex-col gap-1'>
						<span>{t('dialog.formatCells.color')}</span>
						<ColorSelect
							label={t('dialog.formatCells.fontColor')}
							value={draft.fontColor === MIXED ? undefined : draft.fontColor}
							onChange={(c) => onChange({ fontColor: c })}
							theme={theme}
							width={130}
						/>
					</div>
					<Checkbox
						label={t('dialog.formatCells.normalFont')}
						checked={isNormal}
						onChange={(v) => v && onChange(normal)}
					/>
				</div>
			</div>
			<div className='grid grid-cols-[1fr_1fr] gap-3'>
				<GroupBox legend={t('dialog.formatCells.effects')}>
					<div className='flex flex-col gap-1.5 pt-1'>
						<Checkbox
							label={t('dialog.formatCells.strikethrough')}
							checked={draft.strike === true}
							indeterminate={draft.strike === MIXED}
							onChange={(v) => onChange({ strike: v })}
						/>
						<Checkbox
							label={t('dialog.formatCells.superscript')}
							checked={vertAlign === 'superscript'}
							indeterminate={vertAlign === MIXED}
							onChange={(v) => onChange({ vertAlign: v ? 'superscript' : 'baseline' })}
						/>
						<Checkbox
							label={t('dialog.formatCells.subscript')}
							checked={vertAlign === 'subscript'}
							indeterminate={vertAlign === MIXED}
							onChange={(v) => onChange({ vertAlign: v ? 'subscript' : 'baseline' })}
						/>
					</div>
				</GroupBox>
				<GroupBox legend={t('dialog.preview')}>
					<div
						className='flex h-[66px] items-center justify-center overflow-hidden bg-white'
						data-testid='font-preview'
					>
						<span style={previewStyle(draft, theme)}>
							{vertAlign === 'superscript' || vertAlign === 'subscript' ? (
								<>
									Aa
									<span
										style={{
											verticalAlign: vertAlign === 'superscript' ? 'super' : 'sub',
											fontSize: '0.65em',
										}}
									>
										BbCcYyZz
									</span>
								</>
							) : (
								'AaBbCcYyZz'
							)}
						</span>
					</div>
				</GroupBox>
			</div>
			<p className='text-[12px] text-muted-foreground'>{t('dialog.formatCells.trueTypeHint')}</p>
		</div>
	);
}
