import {
	formatHeaderFooter,
	headerFooterText,
	joinHeaderFooter,
	splitHeaderFooter,
} from 'xlsx-viewer-core';
import { useState } from 'react';

import {
	headerFooterPresets,
	type HeaderFooterPreset,
	type PresetLabels,
} from '../../../../engine/print/header-footer-presets';
import type { HeaderFooterKey } from '../../../../engine/print/page-setup-draft';
import { formatNumber, useLocale, useT, type TranslationKey } from '../../../../i18n';
import { Checkbox, FIELD, FIELD_SM } from '../dialog-ui';
import type { DraftTabProps } from './PageTab';

type Variant = 'odd' | 'even' | 'first';
type Part = 'Header' | 'Footer';
type Section = 'left' | 'center' | 'right';

interface PartKeys {
	label: TranslationKey;
	preview: TranslationKey;
	sections: Record<Section, TranslationKey>;
}

const PART_KEYS: Readonly<Record<Part, PartKeys>> = {
	Header: {
		label: 'dialog.pageSetup.header',
		preview: 'dialog.pageSetup.headerPreview',
		sections: {
			left: 'dialog.pageSetup.headerLeft',
			center: 'dialog.pageSetup.headerCenter',
			right: 'dialog.pageSetup.headerRight',
		},
	},
	Footer: {
		label: 'dialog.pageSetup.footer',
		preview: 'dialog.pageSetup.footerPreview',
		sections: {
			left: 'dialog.pageSetup.footerLeft',
			center: 'dialog.pageSetup.footerCenter',
			right: 'dialog.pageSetup.footerRight',
		},
	},
};

const SECTION_PLACEHOLDERS: Readonly<Record<Section, TranslationKey>> = {
	left: 'dialog.pageSetup.sectionLeft',
	center: 'dialog.pageSetup.sectionCenter',
	right: 'dialog.pageSetup.sectionRight',
};

function keyOf(variant: Variant, part: Part): HeaderFooterKey {
	return `${variant}${part}` as HeaderFooterKey;
}

interface PartEditorProps {
	part: Part;
	code: string;
	labels: PresetLabels;
	presets: readonly HeaderFooterPreset[];
	onChange: (code: string) => void;
}

function PartEditor({ part, code, labels, presets, onChange }: PartEditorProps) {
	const t = useT();
	const keys = PART_KEYS[part];
	const sections = splitHeaderFooter(code);
	const preview = headerFooterText(
		formatHeaderFooter(code, { page: 1, pages: 1, sheetName: labels.sheet, fileName: labels.file }),
	);
	const preset = presets.findIndex((p) => p.code === code);
	return (
		<div className='flex flex-col gap-1.5'>
			<div
				className='grid min-h-[34px] grid-cols-3 items-center gap-2 rounded-md border border-border bg-card px-2 py-1 text-[12px]'
				aria-label={t(keys.preview)}
			>
				<span className='truncate whitespace-pre text-left'>{preview.left}</span>
				<span className='truncate whitespace-pre text-center'>{preview.center}</span>
				<span className='truncate whitespace-pre text-right'>{preview.right}</span>
			</div>
			<label className='flex items-center gap-2'>
				<span className='w-[60px] shrink-0'>{t(keys.label)}</span>
				<select
					className={FIELD}
					value={preset >= 0 ? String(preset) : 'custom'}
					onChange={(e) => {
						const p = presets[Number(e.target.value)];
						if (p) {
							onChange(p.code);
						}
					}}
				>
					{preset < 0 && <option value='custom'>{t('dialog.pageSetup.presetCustom')}</option>}
					{presets.map((p, i) => (
						<option key={p.code || 'none'} value={String(i)}>
							{p.label}
						</option>
					))}
				</select>
			</label>
			<div className='grid grid-cols-3 gap-2'>
				{(['left', 'center', 'right'] as const).map((s) => (
					<input
						key={s}
						aria-label={t(keys.sections[s])}
						placeholder={t(SECTION_PLACEHOLDERS[s])}
						className={FIELD_SM}
						value={sections[s]}
						onChange={(e) => onChange(joinHeaderFooter({ ...sections, [s]: e.target.value }))}
					/>
				))}
			</div>
		</div>
	);
}

/** Page Setup > Header/Footer: presets, custom sections, odd / even / first page variants. */
export function HeaderFooterTab({
	draft,
	update,
	sheetName,
	fileName,
}: DraftTabProps & { sheetName: string; fileName: string }) {
	const t = useT();
	const locale = useLocale();
	const [variant, setVariant] = useState<Variant>('odd');
	const labels: PresetLabels = {
		sheet: sheetName,
		file: fileName,
		// &D prints Excel's default short date (en-US), so the presets show that format.
		date: new Date().toLocaleDateString('en-US'),
		page: formatNumber(1, locale),
	};
	const presets = headerFooterPresets(t, labels);
	const variants: Variant[] = ['odd'];
	if (draft.differentOddEven) {
		variants.push('even');
	}
	if (draft.differentFirst) {
		variants.push('first');
	}
	const current = variants.includes(variant) ? variant : 'odd';
	const variantLabel = (v: Variant) =>
		t(
			v === 'first'
				? 'dialog.pageSetup.firstPage'
				: v === 'even'
					? 'dialog.pageSetup.evenPage'
					: variants.length > 1
						? 'dialog.pageSetup.oddPage'
						: 'dialog.pageSetup.allPages',
		);
	const set = (part: Part) => (code: string) =>
		update({ texts: { ...draft.texts, [keyOf(current, part)]: code } });
	return (
		<div className='flex flex-col gap-3'>
			{variants.length > 1 && (
				<div role='radiogroup' aria-label={t('print.pages')} className='flex gap-1'>
					{variants.map((v) => (
						<button
							key={v}
							type='button'
							role='radio'
							aria-checked={current === v}
							className={`rounded-md px-2 py-0.5 text-[12px] ${current === v ? 'bg-primary text-primary-foreground' : 'border border-border hover:bg-accent'}`}
							onClick={() => setVariant(v)}
						>
							{variantLabel(v)}
						</button>
					))}
				</div>
			)}
			<PartEditor
				part='Header'
				code={draft.texts[keyOf(current, 'Header')]}
				labels={labels}
				presets={presets}
				onChange={set('Header')}
			/>
			<PartEditor
				part='Footer'
				code={draft.texts[keyOf(current, 'Footer')]}
				labels={labels}
				presets={presets}
				onChange={set('Footer')}
			/>
			<p className='text-[11.5px] text-muted-foreground'>{t('dialog.pageSetup.codesHelp')}</p>
			<div className='grid grid-cols-2 gap-1.5'>
				<Checkbox
					label={t('dialog.pageSetup.differentOddEven')}
					checked={draft.differentOddEven}
					onChange={(v) => update({ differentOddEven: v })}
				/>
				<Checkbox
					label={t('dialog.pageSetup.differentFirst')}
					checked={draft.differentFirst}
					onChange={(v) => update({ differentFirst: v })}
				/>
				<Checkbox
					label={t('dialog.pageSetup.scaleWithDoc')}
					checked={draft.scaleWithDoc}
					onChange={(v) => update({ scaleWithDoc: v })}
				/>
				<Checkbox
					label={t('dialog.pageSetup.alignWithMargins')}
					checked={draft.alignWithMargins}
					onChange={(v) => update({ alignWithMargins: v })}
				/>
			</div>
		</div>
	);
}
