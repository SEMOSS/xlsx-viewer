import { useLayoutEffect, useState, type RefObject } from 'react';
import { createPortal } from 'react-dom';

import { useT } from '../../../i18n';
import { usePortalHost } from '../ui/Popover';
import type { AssistModel, Completion } from './useFormulaAssist';

const KIND_BADGE: Record<Completion['kind'], { letter: string; color: string }> = {
	function: { letter: 'f', color: 'var(--xlsx-selection-border, #107c41)' },
	name: { letter: 'N', color: '#8e7cc3' },
	table: { letter: 'T', color: '#3c78d8' },
};

interface FormulaAssistPopupProps {
	assist: AssistModel;
	/** The editor element the popup sits under. */
	anchorRef: RefObject<HTMLElement | null>;
	/** Horizontal offset from the anchor's left edge (e.g. the caret position). */
	offsetX?: number;
}

/** Autocomplete list and argument tooltip shown under a formula editor. */
export function FormulaAssistPopup({ assist, anchorRef, offsetX = 0 }: FormulaAssistPopupProps) {
	const t = useT();
	const host = usePortalHost();
	const [pos, setPos] = useState<{ left: number; top: number } | null>(null);
	const visible = assist.completions.length > 0 || assist.hint !== null;

	// The editor grows while typing: re-measure after every render, but only store changes.
	useLayoutEffect(() => {
		const r = anchorRef.current?.getBoundingClientRect();
		if (!r || !visible) {
			return;
		}
		const left = Math.round(Math.min(r.left + offsetX, window.innerWidth - 340));
		const top = Math.round(r.bottom + 2);
		setPos((p) => (p && p.left === left && p.top === top ? p : { left, top }));
	});

	if (!visible || !pos) {
		return null;
	}
	const { hint } = assist;
	const node = (
		<div
			className='xv-formula-assist fixed z-[1200] flex flex-col items-start gap-1'
			style={{ left: pos.left, top: pos.top }}
			data-xv-popover
		>
			{hint && (
				<div
					role='tooltip'
					className='max-w-[560px] rounded-sm border border-[#c6c6c6] bg-[#ffffe1] px-2 py-0.5 text-[12px] whitespace-nowrap text-[#222] shadow-sm'
				>
					<span>{hint.name}(</span>
					{hint.parts.map((p, i) => (
						<span key={i}>
							{i > 0 ? ', ' : ''}
							<span className={p.current ? 'font-bold' : undefined}>{p.label}</span>
							{p.ellipsis ? ', ...' : ''}
						</span>
					))}
					)
				</div>
			)}
			{assist.completions.length > 0 && (
				<div
					role='listbox'
					aria-label={t('formulaAssist.suggestions')}
					className='max-h-[260px] w-[320px] overflow-y-auto rounded-md border border-border bg-popover py-1 text-popover-foreground shadow-lg'
				>
					{assist.completions.map((c, i) => {
						const badge = KIND_BADGE[c.kind];
						const selected = i === assist.highlighted;
						return (
							<div
								key={`${c.kind}:${c.name}`}
								role='option'
								aria-selected={selected}
								ref={selected ? (el) => el?.scrollIntoView({ block: 'nearest' }) : undefined}
								className={`flex cursor-default flex-col px-2.5 py-1 ${selected ? 'bg-accent' : 'hover:bg-accent/60'}`}
								onPointerDown={(e) => {
									// Keep focus in the editor.
									e.preventDefault();
								}}
								onClick={() => assist.accept(i)}
							>
								<span className='flex items-center gap-1.5 text-[12.5px] font-medium'>
									<span
										aria-hidden
										className='inline-flex h-3.5 w-3.5 items-center justify-center rounded-[3px] text-[9px] font-bold text-white italic'
										style={{ background: badge.color }}
									>
										{badge.letter}
									</span>
									{c.name}
								</span>
								{selected && c.description && (
									<span className='line-clamp-3 pl-5 text-[11.5px] text-muted-foreground'>
										{c.description}
									</span>
								)}
							</div>
						);
					})}
				</div>
			)}
		</div>
	);
	return host ? createPortal(node, host) : node;
}
