import {
	ChartArea,
	ChartBar,
	ChartColumn,
	ChartLine,
	ChartPie,
	ChartScatter,
	Image,
	Link,
	StickyNote,
	Table2,
} from 'lucide-react';

import { useT } from '../../../../i18n';
import { useShell } from '../../../shell-context';
import { useViewerStore } from '../../../state/use-store';
import { RibbonButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';
import { useRibbonHost } from '../ribbon-host';

/** Insert: tables, pictures, charts, links and notes. */
export function InsertTab() {
	const t = useT();
	const store = useViewerStore();
	const shell = useShell();
	const host = useRibbonHost();
	const readOnly = store.readOnly;
	const chart = (type: Parameters<NonNullable<typeof host.insertChart>>[0]) => () =>
		host.insertChart?.(type);
	return (
		<>
			<RibbonGroup label={t('group.tables')}>
				<RibbonButton
					icon={<Table2 size={15} />}
					label={t('cmd.table')}
					showLabel
					disabled={readOnly || !host.insertTable}
					onClick={() => host.insertTable?.()}
					shortcut='Ctrl+T'
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.illustrations')}>
				<RibbonButton
					icon={<Image size={15} />}
					label={t('cmd.pictures')}
					showLabel
					disabled={readOnly || !host.insertPicture}
					onClick={() => host.insertPicture?.()}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.charts')}>
				<RibbonButton
					icon={<ChartColumn size={15} />}
					label={t('chart.family.column')}
					showLabel
					disabled={readOnly || !host.insertChart}
					onClick={chart('column')}
				/>
				<RibbonButton
					icon={<ChartBar size={15} />}
					label={t('chart.family.bar')}
					disabled={readOnly || !host.insertChart}
					onClick={chart('bar')}
				/>
				<RibbonButton
					icon={<ChartLine size={15} />}
					label={t('chart.family.line')}
					disabled={readOnly || !host.insertChart}
					onClick={chart('line')}
				/>
				<RibbonButton
					icon={<ChartPie size={15} />}
					label={t('chart.family.pie')}
					disabled={readOnly || !host.insertChart}
					onClick={chart('pie')}
				/>
				<RibbonButton
					icon={<ChartArea size={15} />}
					label={t('chart.family.area')}
					disabled={readOnly || !host.insertChart}
					onClick={chart('area')}
				/>
				<RibbonButton
					icon={<ChartScatter size={15} />}
					label={t('chart.type.scatter')}
					disabled={readOnly || !host.insertChart}
					onClick={chart('scatter')}
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.links')}>
				<RibbonButton
					icon={<Link size={15} />}
					label={t('cmd.link')}
					showLabel
					disabled={readOnly}
					onClick={() => shell.openDialog('hyperlink')}
					shortcut='Ctrl+K'
				/>
			</RibbonGroup>
			<RibbonGroup label={t('group.comments')}>
				<RibbonButton
					icon={<StickyNote size={15} />}
					label={t('cmd.note')}
					showLabel
					disabled={readOnly}
					onClick={() => shell.openDialog('note')}
					shortcut='Shift+F2'
				/>
			</RibbonGroup>
		</>
	);
}
