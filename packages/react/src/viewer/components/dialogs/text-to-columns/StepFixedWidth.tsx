import { useLayoutEffect, useRef, useState, type MouseEvent } from 'react';

import { useT } from '../../../../i18n';
import { FIELD } from '../dialog-ui';
import { addBreak, formatBreaks, parseBreaksText, PREVIEW_LINES } from './wizard-model';

/** Character width used until the monospace font can be measured (and in tests). */
const FALLBACK_CHAR_WIDTH = 8;

const BREAK_LINE =
	'absolute bottom-0 top-[12px] -ml-[4px] flex w-[9px] cursor-col-resize flex-col items-center';
const ARROW = 'border-x-[4px] border-b-[6px] border-x-transparent border-b-foreground';

/**
 * Step 2 for fixed-width data: a ruler and the text with break lines.
 * Click to add a break, double-click a break to remove it, drag it to move
 * it; the break list can also be typed.
 */
export function StepFixedWidth({
	lines,
	breaks,
	onChange,
}: {
	lines: readonly string[];
	breaks: readonly number[];
	onChange: (breaks: number[]) => void;
}) {
	const t = useT();
	const measure = useRef<HTMLSpanElement>(null);
	const area = useRef<HTMLDivElement>(null);
	const [cw, setCw] = useState(FALLBACK_CHAR_WIDTH);
	const [draft, setDraft] = useState<string | null>(null);
	const shown = lines.slice(0, PREVIEW_LINES);
	const width = lines.reduce((n, l) => Math.max(n, l.length), 1);

	useLayoutEffect(() => {
		const w = measure.current?.getBoundingClientRect().width ?? 0;
		if (w > 0) {
			setCw(w / 10);
		}
	}, []);

	const positionAt = (clientX: number): number => {
		const left = area.current?.getBoundingClientRect().left ?? 0;
		return Math.round((clientX - left) / cw);
	};
	const startDrag = (e: MouseEvent, index: number) => {
		e.preventDefault();
		const others = breaks.filter((_, i) => i !== index);
		let current = breaks[index];
		const move = (ev: globalThis.MouseEvent) => {
			const at = Math.max(1, Math.min(width - 1, positionAt(ev.clientX)));
			if (at !== current) {
				current = at;
				onChange([...new Set([...others, at])].sort((a, b) => a - b));
			}
		};
		const up = () => {
			window.removeEventListener('mousemove', move);
			window.removeEventListener('mouseup', up);
		};
		window.addEventListener('mousemove', move);
		window.addEventListener('mouseup', up);
	};
	const remove = (at: number) => onChange(breaks.filter((b) => b !== at));

	const labels = Array.from({ length: Math.floor(width / 10) + 1 }, (_, i) => i * 10);
	const ticks = `repeating-linear-gradient(to right, currentColor 0 1px, transparent 1px ${cw}px)`;
	return (
		<>
			<div>
				<p>{t('dialog.textToColumns.fixedIntro')}</p>
				<p>{t('dialog.textToColumns.arrowsHint')}</p>
				<ul className='ml-5 list-disc'>
					<li>{t('dialog.textToColumns.createBreak')}</li>
					<li>{t('dialog.textToColumns.deleteBreak')}</li>
					<li>{t('dialog.textToColumns.moveBreak')}</li>
				</ul>
			</div>
			<label className='flex items-center gap-2'>
				<span className='whitespace-nowrap'>{t('dialog.textToColumns.breakPositions')}</span>
				<input
					className={FIELD}
					value={draft ?? formatBreaks(breaks)}
					placeholder={t('dialog.textToColumns.breakExample')}
					onChange={(e) => {
						setDraft(e.target.value);
						onChange(parseBreaksText(e.target.value).filter((b) => b < width));
					}}
					onBlur={() => setDraft(null)}
				/>
			</label>
			<div className='flex min-w-0 flex-col gap-1'>
				<span>{t('dialog.textToColumns.dataPreview')}</span>
				<div className='h-[150px] overflow-auto rounded-md border border-input bg-card px-2 py-1'>
					<span
						ref={measure}
						aria-hidden
						className='invisible absolute whitespace-pre font-mono text-[12px]'
					>
						0000000000
					</span>
					<div
						ref={area}
						data-testid='fixed-width-area'
						className='relative cursor-text font-mono text-[12px]'
						style={{ width: (width + 2) * cw }}
						onClick={(e) => onChange(addBreak(breaks, positionAt(e.clientX), width))}
						onDoubleClick={(e) => remove(positionAt(e.clientX))}
					>
						<div aria-hidden className='relative h-[22px] text-[10px] text-muted-foreground'>
							{labels.map((n) => (
								<span key={n} className='absolute top-0' style={{ left: n * cw }}>
									{n}
								</span>
							))}
							<div
								className='absolute bottom-0 left-0 right-0 h-[5px] text-border'
								style={{ backgroundImage: ticks }}
							/>
						</div>
						{shown.map((line, i) => (
							<div key={i} className='whitespace-pre leading-[18px]'>
								{line || ' '}
							</div>
						))}
						{breaks.map((b, i) => (
							<div
								key={b}
								data-break={b}
								aria-hidden
								className={BREAK_LINE}
								style={{ left: b * cw }}
								onMouseDown={(e) => startDrag(e, i)}
								onDoubleClick={(e) => {
									e.stopPropagation();
									remove(b);
								}}
							>
								<span className={ARROW} />
								<span className='w-px flex-1 bg-foreground' />
							</div>
						))}
					</div>
				</div>
			</div>
		</>
	);
}
