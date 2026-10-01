import { useRef, type KeyboardEvent, type ReactNode } from 'react';

import { cx } from './dialog-cx';

export interface TabDef<T extends string> {
	id: T;
	label: string;
}

/** Ids linking a tab to its panel. */
export function tabIds(base: string, id: string): { tab: string; panel: string } {
	return { tab: `${base}-tab-${id}`, panel: `${base}-panel-${id}` };
}

/** Accessible tab strip (arrow keys, Home / End move between tabs). */
export function TabStrip<T extends string>({
	tabs,
	value,
	onChange,
	label,
	idBase,
}: {
	tabs: readonly TabDef<T>[];
	value: T;
	onChange: (id: T) => void;
	label: string;
	idBase: string;
}) {
	const refs = useRef<Array<HTMLButtonElement | null>>([]);
	const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
		const i = tabs.findIndex((t) => t.id === value);
		let next = -1;
		if (e.key === 'ArrowRight') {
			next = (i + 1) % tabs.length;
		} else if (e.key === 'ArrowLeft') {
			next = (i - 1 + tabs.length) % tabs.length;
		} else if (e.key === 'Home') {
			next = 0;
		} else if (e.key === 'End') {
			next = tabs.length - 1;
		}
		if (next >= 0) {
			e.preventDefault();
			onChange(tabs[next].id);
			refs.current[next]?.focus();
		}
	};
	return (
		<div
			role='tablist'
			aria-label={label}
			className='flex gap-0.5 border-b border-border'
			onKeyDown={onKeyDown}
		>
			{tabs.map((t, i) => {
				const ids = tabIds(idBase, t.id);
				const selected = t.id === value;
				return (
					<button
						key={t.id}
						ref={(el) => {
							refs.current[i] = el;
						}}
						id={ids.tab}
						type='button'
						role='tab'
						aria-selected={selected}
						aria-controls={ids.panel}
						tabIndex={selected ? 0 : -1}
						className={cx(
							'-mb-px border-b-2 px-3 py-1.5 text-[13px]',
							selected
								? 'border-primary font-semibold text-foreground'
								: 'border-transparent text-muted-foreground hover:text-foreground',
						)}
						onClick={() => onChange(t.id)}
					>
						{t.label}
					</button>
				);
			})}
		</div>
	);
}

/** Panel of a {@link TabStrip} tab. */
export function TabPanel({
	idBase,
	id,
	children,
	className,
}: {
	idBase: string;
	id: string;
	children: ReactNode;
	className?: string;
}) {
	const ids = tabIds(idBase, id);
	return (
		<div role='tabpanel' id={ids.panel} aria-labelledby={ids.tab} className={className}>
			{children}
		</div>
	);
}
