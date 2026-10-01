import { ArrowLeftToLine, ArrowRightFromLine, Calculator, Eraser, Eye, Tag } from 'lucide-react';

import { useT } from '../../../../i18n';
import { removeArrows, traceDependents, tracePrecedents } from '../../../commands/audit-commands';
import { insertNameInFormula, toggleShowFormulas } from '../../../commands/misc-commands';
import { useShell } from '../../../shell-context';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import { MenuButton, RibbonButton, SplitButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';
import { FunctionLibraryGroup } from './FunctionLibraryGroup';

/** Formulas: function library, defined names, formula auditing and calculation. */
export function FormulasTab() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const manual = useStoreSelector((s) => s.ui.manualCalc);
	const showFormulas = useStoreSelector(
		(s, st) => (void s.docVersion, st.workbook.sheets[s.ui.activeSheet].view.showFormulas),
	);
	const names = useStoreSelector(
		(s, st) => (
			void s.docVersion,
			st.workbook.definedNames
				.filter(
					(n) =>
						!n.hidden &&
						!n.builtin &&
						(n.localSheet === undefined || n.localSheet === s.ui.activeSheet),
				)
				.map((n) => n.name)
		),
		(a, b) => a.length === b.length && a.every((x, i) => x === b[i]),
	);
	return (
		<>
			<FunctionLibraryGroup />
			<RibbonGroup label={t('group.definedNames')}>
				<RibbonButton
					icon={<Tag size={15} />}
					label={t('cmd.nameManager')}
					showLabel
					onClick={() => shell.openDialog('nameManager')}
				/>
				<RibbonButton
					icon={<Tag size={15} />}
					label={t('cmd.defineName')}
					showLabel
					disabled={store.readOnly}
					onClick={() => shell.openDialog('defineName')}
				/>
				<MenuButton
					label={t('cmd.useInFormula')}
					disabled={store.readOnly}
					items={
						names.length
							? names.map((n) => ({ label: n, onSelect: () => insertNameInFormula(store, n) }))
							: [{ label: t('cmd.noNames'), disabled: true }]
					}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.formulaAuditing')}>
				<RibbonButton
					icon={<ArrowLeftToLine size={15} />}
					label={t('cmd.tracePrecedents')}
					showLabel
					onClick={() => tracePrecedents(store)}
				/>
				<RibbonButton
					icon={<ArrowRightFromLine size={15} />}
					label={t('cmd.traceDependents')}
					showLabel
					onClick={() => traceDependents(store)}
				/>
				<SplitButton
					icon={<Eraser size={15} />}
					label={t('cmd.removeArrows')}
					showLabel
					onClick={() => removeArrows(store)}
					items={[
						{ label: t('cmd.removeArrows'), onSelect: () => removeArrows(store) },
						{
							label: t('cmd.removePrecedentArrows'),
							onSelect: () => removeArrows(store, 'precedent'),
						},
						{
							label: t('cmd.removeDependentArrows'),
							onSelect: () => removeArrows(store, 'dependent'),
						},
					]}
				/>
				<RibbonButton
					icon={<Eye size={15} />}
					label={t('cmd.showFormulas')}
					showLabel
					pressed={showFormulas}
					onClick={() => toggleShowFormulas(store)}
					shortcut='Ctrl+`'
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.calculation')}>
				<MenuButton
					icon={<Calculator size={15} />}
					label={t('cmd.calcOptions')}
					items={[
						{
							label: t('cmd.automatic'),
							checked: !manual,
							onSelect: () => store.setUi({ manualCalc: false }),
						},
						{
							label: t('cmd.manual'),
							checked: manual,
							onSelect: () => store.setUi({ manualCalc: true }),
						},
					]}
				/>
				<RibbonButton
					icon={<Calculator size={15} />}
					label={t('cmd.calculateNow')}
					showLabel
					onClick={() => store.recalculate()}
					shortcut='F9'
				/>
			</RibbonGroup>
		</>
	);
}
