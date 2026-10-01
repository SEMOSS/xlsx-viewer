import { useEffect, useRef, type ReactNode } from 'react';

import { cx } from './dialog-cx';
import { FIELD } from './dialog-fields';

/** Check box with its label; `indeterminate` shows the mixed state. */
export function Checkbox({
	label,
	checked,
	onChange,
	disabled,
	indeterminate,
	className,
}: {
	label: ReactNode;
	checked: boolean;
	onChange: (checked: boolean) => void;
	disabled?: boolean;
	indeterminate?: boolean;
	className?: string;
}) {
	const ref = useRef<HTMLInputElement>(null);
	useEffect(() => {
		if (ref.current) {
			ref.current.indeterminate = Boolean(indeterminate);
		}
	}, [indeterminate]);
	return (
		<label className={cx('flex items-center gap-1.5', disabled && 'opacity-50', className)}>
			<input
				ref={ref}
				type='checkbox'
				className='accent-[var(--xlsx-primary)]'
				checked={checked}
				disabled={disabled}
				onChange={(e) => onChange(e.target.checked)}
			/>
			<span>{label}</span>
		</label>
	);
}

/** Radio button with its label. */
export function Radio({
	name,
	label,
	checked,
	onChange,
	disabled,
	className,
}: {
	name: string;
	label: ReactNode;
	checked: boolean;
	onChange: () => void;
	disabled?: boolean;
	className?: string;
}) {
	return (
		<label className={cx('flex items-center gap-1.5', disabled && 'opacity-50', className)}>
			<input
				type='radio'
				name={name}
				className='accent-[var(--xlsx-primary)]'
				checked={checked}
				disabled={disabled}
				onChange={onChange}
			/>
			<span>{label}</span>
		</label>
	);
}

/** Titled frame around related controls (Windows "group box"). */
export function GroupBox({
	legend,
	children,
	className,
}: {
	legend: string;
	children: ReactNode;
	className?: string;
}) {
	return (
		<fieldset
			className={cx('min-w-0 rounded-md border border-border px-3 pb-2.5 pt-0.5', className)}
		>
			<legend className='px-1 text-[12px] font-semibold text-muted-foreground'>{legend}</legend>
			{children}
		</fieldset>
	);
}

/** A label above its control; the control is labelled by wrapping. */
export function LabeledField({
	label,
	children,
	className,
}: {
	label: string;
	children: ReactNode;
	className?: string;
}) {
	return (
		<label className={cx('flex min-w-0 flex-col gap-1', className)}>
			<span>{label}</span>
			{children}
		</label>
	);
}

export interface SelectOption<T extends string> {
	value: T;
	label: string;
}

/** Native drop-down list. */
export function Select<T extends string>({
	value,
	options,
	onChange,
	label,
	disabled,
	className,
}: {
	value: T;
	options: readonly SelectOption<T>[];
	onChange: (value: T) => void;
	/** Accessible name when there is no wrapping label. */
	label?: string;
	disabled?: boolean;
	className?: string;
}) {
	return (
		<select
			aria-label={label}
			className={cx(FIELD, className)}
			value={value}
			disabled={disabled}
			onChange={(e) => onChange(e.target.value as T)}
		>
			{options.map((o) => (
				<option key={o.value} value={o.value}>
					{o.label}
				</option>
			))}
		</select>
	);
}

/** Inline error / hint line. */
export function DialogMessage({
	children,
	tone = 'error',
}: {
	children: ReactNode;
	tone?: 'error' | 'info';
}) {
	return (
		<p
			role={tone === 'error' ? 'alert' : 'status'}
			className={cx('text-[12px]', tone === 'error' ? 'text-destructive' : 'text-muted-foreground')}
		>
			{children}
		</p>
	);
}
