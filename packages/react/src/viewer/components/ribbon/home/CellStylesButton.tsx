import { SwatchBook } from 'lucide-react';

import { useT } from '../../../../i18n';
import { applyCellStyleCmd } from '../../../commands/cell-style-commands';
import { useViewerStore } from '../../../state/use-store';
import { CellStyleGallery } from '../../styles/CellStyleGallery';
import { MenuButton } from '../../ui/buttons';

/** Home > Styles > Cell Styles. */
export function CellStylesButton() {
	const t = useT();
	const store = useViewerStore();
	return (
		<MenuButton
			icon={<SwatchBook size={15} />}
			label={t('cmd.cellStyles')}
			disabled={store.readOnly}
			renderPanel={(close) => (
				<CellStyleGallery
					onPick={(name) => {
						applyCellStyleCmd(store, name);
						close();
					}}
				/>
			)}
		/>
	);
}
