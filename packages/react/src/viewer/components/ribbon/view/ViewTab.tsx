import { Grid3x3, PanelTop, Snowflake, SquareFunction, ZoomIn } from 'lucide-react';

import { useT } from '../../../../i18n';
import { freezePanes, setZoom, toggleSheetViewFlag } from '../../../commands/view-commands';
import { useShell } from '../../../shell-context';
import { useStoreSelector, useViewerStore } from '../../../state/use-store';
import { MenuButton, RibbonButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';
import { useRibbonHost } from '../ribbon-host';

/** View: show options, zoom, freeze panes and the theme switch. */
export function ViewTab() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const host = useRibbonHost();
	const view = useStoreSelector(
		(s, st) => (void s.docVersion, st.workbook.sheets[s.ui.activeSheet].view),
		(a, b) => a === b,
	);
	const frozen = view.pane?.state === 'frozen';
	return (
		<>
			<RibbonGroup label={t('group.show')}>
				<RibbonButton
					icon={<Grid3x3 size={15} />}
					label={t('cmd.gridlines')}
					showLabel
					pressed={view.showGridLines}
					onClick={() => toggleSheetViewFlag(store, 'showGridLines')}
				/>
				<RibbonButton
					icon={<PanelTop size={15} />}
					label={t('cmd.headings')}
					showLabel
					pressed={view.showRowColHeaders}
					onClick={() => toggleSheetViewFlag(store, 'showRowColHeaders')}
				/>
				<RibbonButton
					icon={<SquareFunction size={15} />}
					label={t('cmd.formulaBar')}
					showLabel
					pressed={host.formulaBarVisible}
					onClick={() => host.toggleFormulaBar?.()}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.zoom')}>
				<RibbonButton
					icon={<ZoomIn size={15} />}
					label={t('cmd.zoom')}
					showLabel
					onClick={() => shell.openDialog('zoom')}
				/>
				<RibbonButton label={t('cmd.zoom100')} showLabel onClick={() => setZoom(store, 100)} />
				<RibbonButton
					label={t('cmd.zoomSelection')}
					showLabel
					onClick={() => host.zoomToSelection?.()}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.window')}>
				<MenuButton
					icon={<Snowflake size={15} />}
					label={t('cmd.freezePanes')}
					disabled={store.readOnly}
					items={[
						{
							label: frozen ? t('cmd.unfreezePanes') : t('cmd.freezePanes'),
							onSelect: () =>
								freezePanes(store, frozen ? 'none' : 'panes', host.topLeftVisible?.()),
						},
						{
							label: t('cmd.freezeTopRow'),
							onSelect: () => freezePanes(store, 'topRow', host.topLeftVisible?.()),
						},
						{
							label: t('cmd.freezeFirstColumn'),
							onSelect: () => freezePanes(store, 'firstColumn', host.topLeftVisible?.()),
						},
					]}
				/>
			</RibbonGroup>
			{shell.toggleTheme && (
				<RibbonGroup label={t('group.appearance')}>
					<RibbonButton
						label={shell.themeMode === 'dark' ? t('cmd.lightMode') : t('cmd.darkMode')}
						showLabel
						onClick={shell.toggleTheme}
					/>
				</RibbonGroup>
			)}
		</>
	);
}
