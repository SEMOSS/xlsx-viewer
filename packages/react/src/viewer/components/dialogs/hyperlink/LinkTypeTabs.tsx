import { FileSpreadsheet, Globe, Mail } from 'lucide-react';
import { useRef, type KeyboardEvent, type ReactNode } from 'react';

import { useT, type TranslationKey } from '../../../../i18n';
import type { LinkKind } from '../../../commands/hyperlink-dialog-commands';
import { tabIds } from '../dialog-ui';

const TABS: ReadonlyArray<{ id: LinkKind; labelKey: TranslationKey; icon: ReactNode }> = [
	{
		id: 'file',
		labelKey: 'dialog.hyperlink.fileOrWebPage',
		icon: <Globe size={22} aria-hidden='true' />,
	},
	{
		id: 'place',
		labelKey: 'dialog.hyperlink.placeInDocument',
		icon: <FileSpreadsheet size={22} aria-hidden='true' />,
	},
	{
		id: 'email',
		labelKey: 'dialog.hyperlink.emailAddress',
		icon: <Mail size={22} aria-hidden='true' />,
	},
];

const TAB =
	'flex flex-col items-center gap-1 rounded-md border px-1.5 py-2 text-center text-[12px]';
const TAB_ON = 'border-primary bg-card font-semibold text-foreground';
const TAB_OFF = 'border-transparent text-muted-foreground hover:bg-accent hover:text-foreground';

/**
 * The Insert Hyperlink "Link to:" bar: a vertical tab list (arrow keys,
 * Home / End) whose panels are rendered by the dialog with `tabIds`.
 */
export function LinkTypeTabs({
	value,
	onChange,
	idBase,
	labelledBy,
}: {
	value: LinkKind;
	onChange: (kind: LinkKind) => void;
	idBase: string;
	labelledBy: string;
}) {
	const t = useT();
	const refs = useRef<Array<HTMLButtonElement | null>>([]);
	const onKeyDown = (e: KeyboardEvent<HTMLDivElement>) => {
		const i = TABS.findIndex((tab) => tab.id === value);
		const moves: Record<string, number> = {
			ArrowDown: (i + 1) % TABS.length,
			ArrowUp: (i - 1 + TABS.length) % TABS.length,
			Home: 0,
			End: TABS.length - 1,
		};
		const next = moves[e.key] ?? -1;
		if (next >= 0) {
			e.preventDefault();
			onChange(TABS[next].id);
			refs.current[next]?.focus();
		}
	};
	return (
		<div
			role='tablist'
			aria-orientation='vertical'
			aria-labelledby={labelledBy}
			className='flex w-[118px] flex-col gap-1 rounded-md border border-border bg-muted p-1'
			onKeyDown={onKeyDown}
		>
			{TABS.map((tab, i) => {
				const ids = tabIds(idBase, tab.id);
				const selected = tab.id === value;
				return (
					<button
						key={tab.id}
						ref={(el) => {
							refs.current[i] = el;
						}}
						id={ids.tab}
						type='button'
						role='tab'
						aria-selected={selected}
						aria-controls={ids.panel}
						tabIndex={selected ? 0 : -1}
						className={`${TAB} ${selected ? TAB_ON : TAB_OFF}`}
						onClick={() => onChange(tab.id)}
					>
						{tab.icon}
						<span>{t(tab.labelKey)}</span>
					</button>
				);
			})}
		</div>
	);
}
