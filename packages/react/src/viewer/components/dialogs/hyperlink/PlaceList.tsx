import { ChevronDown } from 'lucide-react';
import { useEffect, useId, useRef, type KeyboardEvent } from 'react';

import { useT, type TranslationKey } from '../../../../i18n';
import type { PlaceOption } from '../../../commands/hyperlink-dialog-commands';

const GROUPS: ReadonlyArray<{ id: PlaceOption['group']; labelKey: TranslationKey }> = [
	{ id: 'sheet', labelKey: 'dialog.hyperlink.cellReferenceGroup' },
	{ id: 'name', labelKey: 'dialog.hyperlink.definedNamesGroup' },
];

const LIST =
	'h-[156px] overflow-y-auto rounded-md border border-input bg-card py-0.5 text-[13px] ' +
	'outline-none focus:border-primary';
const OPTION = 'cursor-default select-none py-[3px] pl-7 pr-2';
const OPTION_ON = 'bg-primary text-primary-foreground';
const OPTION_OFF = 'hover:bg-accent';

/**
 * "Or select a place in this document": sheets under "Cell Reference" and
 * defined names under "Defined Names", as one list box with groups (arrow
 * keys, Home / End move across both groups; Enter is left to the dialog).
 */
export function PlaceList({
	places,
	value,
	onChange,
	label,
}: {
	places: readonly PlaceOption[];
	value: string;
	onChange: (key: string) => void;
	label: string;
}) {
	const t = useT();
	const id = useId();
	const listRef = useRef<HTMLDivElement>(null);
	const index = places.findIndex((p) => p.key === value);
	const optionId = (i: number) => `${id}-opt-${i}`;

	useEffect(() => {
		if (index >= 0) {
			listRef.current
				?.querySelector(`[id="${optionId(index)}"]`)
				?.scrollIntoView?.({ block: 'nearest' });
		}
		// eslint-disable-next-line react-hooks/exhaustive-deps
	}, [index]);

	const move = (to: number) => {
		if (places.length > 0) {
			onChange(places[Math.max(0, Math.min(places.length - 1, to))].key);
		}
	};

	const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
		const moves: Record<string, number> = {
			ArrowDown: index + 1,
			ArrowUp: index < 0 ? 0 : index - 1,
			Home: 0,
			End: places.length - 1,
			PageDown: index + 6,
			PageUp: index - 6,
		};
		const to = moves[e.key];
		if (to !== undefined) {
			e.preventDefault();
			move(to);
		}
	};

	return (
		<div
			ref={listRef}
			role='listbox'
			tabIndex={0}
			aria-label={label}
			aria-activedescendant={index >= 0 ? optionId(index) : undefined}
			className={LIST}
			onKeyDown={onKeyDown}
		>
			{GROUPS.map((group) => {
				const items = places.filter((p) => p.group === group.id);
				if (items.length === 0) {
					return null;
				}
				return (
					<div key={group.id} role='group' aria-labelledby={`${id}-${group.id}`}>
						<div
							id={`${id}-${group.id}`}
							role='presentation'
							className='flex items-center gap-1 px-2 py-[3px] font-semibold'
						>
							<ChevronDown size={12} aria-hidden='true' />
							{t(group.labelKey)}
						</div>
						{items.map((item) => {
							const i = places.indexOf(item);
							const selected = item.key === value;
							return (
								<div
									key={item.key}
									id={optionId(i)}
									role='option'
									aria-selected={selected}
									className={`${OPTION} ${selected ? OPTION_ON : OPTION_OFF}`}
									onClick={() => {
										listRef.current?.focus();
										onChange(item.key);
									}}
								>
									{item.label}
								</div>
							);
						})}
					</div>
				);
			})}
		</div>
	);
}
