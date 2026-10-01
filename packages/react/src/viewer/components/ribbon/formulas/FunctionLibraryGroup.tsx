import type { FunctionCategory } from 'xlsx-viewer-core';
import { listFunctions } from 'xlsx-viewer-core';
import {
	Calendar,
	CircleDollarSign,
	Clock3,
	Ellipsis,
	Search,
	Sigma,
	SquareFunction,
	ToggleLeft,
	Type,
	Variable,
} from 'lucide-react';
import type { ReactNode } from 'react';

import { FUNCTION_CATEGORY_KEYS, useT } from '../../../../i18n';
import { recentFunctions, insertFunctionCall } from '../../../commands/insert-function-commands';
import { autoSum } from '../../../commands/misc-commands';
import { useViewerStore } from '../../../state/use-store';
import { useShell } from '../../../shell-context';
import { MenuList, type MenuItem } from '../../ui/Menu';
import { MenuButton, RibbonButton, SplitButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';

const LIST_CLASS = 'max-h-[60vh] overflow-y-auto';

const CATEGORY_BUTTONS: Array<{ category: FunctionCategory; icon: ReactNode }> = [
	{ category: 'Financial', icon: <CircleDollarSign size={15} /> },
	{ category: 'Logical', icon: <ToggleLeft size={15} /> },
	{ category: 'Text', icon: <Type size={15} /> },
	{ category: 'Date & Time', icon: <Calendar size={15} /> },
	{ category: 'Lookup & Reference', icon: <Search size={15} /> },
	{ category: 'Math & Trig', icon: <Variable size={15} /> },
];

const MORE: FunctionCategory[] = [
	'Statistical',
	'Engineering',
	'Cube',
	'Information',
	'Compatibility',
	'Web',
];

/** Formulas > Function Library: Insert Function, AutoSum, Recently Used and the category menus. */
export function FunctionLibraryGroup() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const ro = store.readOnly;
	const items = (category: FunctionCategory): MenuItem[] =>
		listFunctions(category).map((f) => ({
			label: f.name,
			title: f.description,
			onSelect: () => insertFunctionCall(store, f.name),
		}));
	const more = (): MenuItem[] => [
		...MORE.map((c) => ({ label: t(FUNCTION_CATEGORY_KEYS[c]), items: items(c) })).filter(
			(m) => m.items.length > 0,
		),
		{ separator: true },
		{ label: t('cmd.insertFunction'), onSelect: () => shell.openDialog('insertFunction') },
	];
	const recent = (): MenuItem[] => {
		const names = recentFunctions();
		return names.length
			? names.map((n) => ({ label: n, onSelect: () => insertFunctionCall(store, n) }))
			: [{ label: t('formula.noRecent'), disabled: true }];
	};
	return (
		<RibbonGroup label={t('group.functionLibrary')}>
			<RibbonButton
				icon={<SquareFunction size={15} />}
				label={t('cmd.insertFunction')}
				showLabel
				disabled={ro}
				onClick={() => shell.openDialog('insertFunction')}
			/>
			<SplitButton
				icon={<Sigma size={15} />}
				label={t('cmd.autoSum')}
				showLabel
				disabled={ro}
				onClick={() => autoSum(store)}
				items={[
					{ label: t('cmd.sum'), onSelect: () => autoSum(store, 'SUM') },
					{ label: t('cmd.average'), onSelect: () => autoSum(store, 'AVERAGE') },
					{ label: t('cmd.countNumbers'), onSelect: () => autoSum(store, 'COUNT') },
					{ label: t('cmd.max'), onSelect: () => autoSum(store, 'MAX') },
					{ label: t('cmd.min'), onSelect: () => autoSum(store, 'MIN') },
					{ separator: true },
					{ label: t('formula.moreFunctions'), onSelect: () => shell.openDialog('insertFunction') },
				]}
			/>
			<MenuButton
				icon={<Clock3 size={15} />}
				label={t('formula.recentlyUsed')}
				disabled={ro}
				renderPanel={(close) => (
					<MenuList items={recent()} onClose={close} className={LIST_CLASS} />
				)}
			/>
			{CATEGORY_BUTTONS.map(({ category, icon }) => (
				<MenuButton
					key={category}
					icon={icon}
					label={t(FUNCTION_CATEGORY_KEYS[category])}
					disabled={ro}
					renderPanel={(close) => (
						<MenuList items={items(category)} onClose={close} className={LIST_CLASS} />
					)}
				/>
			))}
			<MenuButton
				icon={<Ellipsis size={15} />}
				label={t('formula.moreFunctions')}
				disabled={ro}
				renderPanel={(close) => <MenuList items={more()} onClose={close} />}
			/>
		</RibbonGroup>
	);
}
