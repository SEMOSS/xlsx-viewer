import { Table2 } from 'lucide-react';

import { useT } from '../../../../i18n';
import { formatAsTable } from '../../../commands/table-commands';
import { useViewerStore } from '../../../state/use-store';
import { TableStyleGallery } from '../../tables/TableStyleGallery';
import { MenuButton } from '../../ui/buttons';

/** Home > Styles > Format as Table: pick a style for the table at the cursor, or make one. */
export function FormatAsTableButton() {
	const t = useT();
	const store = useViewerStore();
	return (
		<MenuButton
			icon={<Table2 size={15} />}
			label={t('cmd.formatAsTable')}
			disabled={store.readOnly}
			renderPanel={(close) => (
				<TableStyleGallery
					onPick={(name) => {
						if (name) {
							formatAsTable(store, name);
						}
						close();
					}}
				/>
			)}
		/>
	);
}
