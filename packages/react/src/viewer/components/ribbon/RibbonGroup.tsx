import type { ReactNode } from 'react';

/** A labelled group of ribbon controls. */
export function RibbonGroup({ label, children }: { label: string; children: ReactNode }) {
	return (
		<section
			className='xv-rgroup flex shrink-0 flex-col items-stretch justify-between border-r border-border px-2 last:border-r-0'
			aria-label={label}
		>
			<div className='flex min-h-[30px] items-center gap-1'>{children}</div>
			<div className='pt-0.5 text-center text-[10.5px] leading-[13px] text-muted-foreground select-none'>
				{label}
			</div>
		</section>
	);
}

/** Thin vertical separator between controls inside a group. */
export function RibbonDivider() {
	return <div className='mx-0.5 h-5 w-px bg-border' aria-hidden />;
}
