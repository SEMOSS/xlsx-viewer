import type { ColorSpec, Theme } from 'xlsx-viewer-core';
import { resolveColorCss } from 'xlsx-viewer-core';
import { ChevronDown } from 'lucide-react';
import { useRef, useState } from 'react';

import { useT } from '../../../i18n';
import { ColorPicker } from '../ui/ColorPicker';
import { Popover } from '../ui/Popover';

interface ColorSelectProps {
	/** Accessible name, e.g. "Color". */
	label: string;
	/** Current color; `undefined` is Automatic / No Color. */
	value: ColorSpec | undefined;
	onChange: (color: ColorSpec | undefined) => void;
	theme: Theme;
	/** Label of the "none" entry (default Automatic, for fonts and borders; No Color for fills). */
	noneLabel?: string;
	/** Swatch shown for the automatic color. */
	automaticColor?: string;
	disabled?: boolean;
	width?: number;
}

/** Drop-down color button for dialogs: a swatch that opens the Office palette. */
export function ColorSelect({
	label,
	value,
	onChange,
	theme,
	noneLabel,
	automaticColor = '#000000',
	disabled,
	width = 120,
}: ColorSelectProps) {
	const t = useT();
	const none = noneLabel ?? t('color.automatic');
	const ref = useRef<HTMLButtonElement>(null);
	const [open, setOpen] = useState(false);
	const css = resolveColorCss(value, { theme });
	return (
		<>
			<button
				ref={ref}
				type='button'
				aria-label={label}
				aria-haspopup='dialog'
				aria-expanded={open}
				disabled={disabled}
				className='inline-flex h-7 items-center gap-1.5 rounded-md border border-input bg-card px-1.5 text-[12px] hover:border-primary disabled:opacity-50'
				style={{ width }}
				onClick={() => setOpen((v) => !v)}
			>
				<span
					className='h-4 flex-1 rounded-[2px] border border-black/15'
					style={{ backgroundColor: css ?? automaticColor }}
					data-testid='color-swatch'
				/>
				{!css && <span className='shrink-0 text-muted-foreground'>{none}</span>}
				<ChevronDown size={13} className='shrink-0 text-muted-foreground' />
			</button>
			<Popover open={open} onClose={() => setOpen(false)} anchorRef={ref} ariaLabel={label}>
				<ColorPicker
					theme={theme}
					noneLabel={none}
					close={() => setOpen(false)}
					onPick={(c) => onChange(c ?? undefined)}
				/>
			</Popover>
		</>
	);
}
