import {
	AlignCenter,
	AlignEndVertical,
	AlignLeft,
	AlignRight,
	AlignStartVertical,
	AlignCenterVertical,
	ListIndentDecrease,
	ListIndentIncrease,
	RotateCcw,
	TableCellsMerge,
	WrapText,
} from 'lucide-react';

import { useT } from '../../../../i18n';
import {
	changeIndent,
	mergeCells,
	setTextRotation,
	toggleHorizontal,
	toggleVertical,
	toggleWrapText,
} from '../../../commands/format-commands';
import { useActiveFormat } from '../../../hooks/useActiveFormat';
import { useViewerStore } from '../../../state/use-store';
import { ControlCluster, MenuButton, RibbonButton, SplitButton } from '../../ui/buttons';
import { RibbonGroup } from '../RibbonGroup';

export function AlignmentGroup() {
	const t = useT();
	const store = useViewerStore();
	const { style, merged } = useActiveFormat();
	const a = style.alignment;
	const readOnly = store.readOnly;
	const v = a.vertical ?? 'bottom';
	return (
		<RibbonGroup label={t('group.alignment')}>
			<ControlCluster>
				{/* Vertical alignment: top / middle / bottom (lucide's "vertical" icons are rotated). */}
				<RibbonButton
					icon={<AlignStartVertical size={15} className='rotate-90' />}
					label={t('cmd.alignTop')}
					pressed={v === 'top'}
					disabled={readOnly}
					onClick={() => toggleVertical(store, 'top')}
				/>
				<RibbonButton
					icon={<AlignCenterVertical size={15} className='rotate-90' />}
					label={t('cmd.alignMiddle')}
					pressed={v === 'center'}
					disabled={readOnly}
					onClick={() => toggleVertical(store, 'center')}
				/>
				<RibbonButton
					icon={<AlignEndVertical size={15} className='rotate-90' />}
					label={t('cmd.alignBottom')}
					pressed={v === 'bottom'}
					disabled={readOnly}
					onClick={() => toggleVertical(store, 'bottom')}
				/>
			</ControlCluster>
			<ControlCluster>
				<RibbonButton
					icon={<AlignLeft size={15} />}
					label={t('cmd.alignLeft')}
					pressed={a.horizontal === 'left'}
					disabled={readOnly}
					onClick={() => toggleHorizontal(store, 'left')}
				/>
				<RibbonButton
					icon={<AlignCenter size={15} />}
					label={t('cmd.alignCenter')}
					pressed={a.horizontal === 'center'}
					disabled={readOnly}
					onClick={() => toggleHorizontal(store, 'center')}
				/>
				<RibbonButton
					icon={<AlignRight size={15} />}
					label={t('cmd.alignRight')}
					pressed={a.horizontal === 'right'}
					disabled={readOnly}
					onClick={() => toggleHorizontal(store, 'right')}
				/>
			</ControlCluster>
			<ControlCluster>
				<RibbonButton
					icon={<ListIndentDecrease size={15} />}
					label={t('cmd.decreaseIndent')}
					disabled={readOnly}
					onClick={() => changeIndent(store, -1)}
				/>
				<RibbonButton
					icon={<ListIndentIncrease size={15} />}
					label={t('cmd.increaseIndent')}
					disabled={readOnly}
					onClick={() => changeIndent(store, 1)}
				/>
			</ControlCluster>
			<MenuButton
				icon={<RotateCcw size={15} />}
				label={t('cmd.orientation')}
				showLabel={false}
				disabled={readOnly}
				items={[
					{ label: t('cmd.angleCounterclockwise'), onSelect: () => setTextRotation(store, 45) },
					{ label: t('cmd.angleClockwise'), onSelect: () => setTextRotation(store, 135) },
					{ label: t('cmd.verticalText'), onSelect: () => setTextRotation(store, 255) },
					{ label: t('cmd.rotateTextUp'), onSelect: () => setTextRotation(store, 90) },
					{ label: t('cmd.rotateTextDown'), onSelect: () => setTextRotation(store, 180) },
					{ separator: true },
					{ label: t('cmd.horizontalText'), onSelect: () => setTextRotation(store, undefined) },
				]}
			/>
			<RibbonButton
				icon={<WrapText size={15} />}
				label={t('cmd.wrapText')}
				showLabel
				pressed={!!a.wrapText}
				disabled={readOnly}
				onClick={() => toggleWrapText(store)}
			/>
			<SplitButton
				icon={<TableCellsMerge size={15} />}
				label={t('cmd.mergeCenter')}
				showLabel
				pressed={merged}
				disabled={readOnly}
				onClick={() => mergeCells(store, 'center')}
				items={[
					{ label: t('cmd.mergeCenter'), onSelect: () => mergeCells(store, 'center') },
					{ label: t('cmd.mergeAcross'), onSelect: () => mergeCells(store, 'across') },
					{ label: t('cmd.mergeCells'), onSelect: () => mergeCells(store, 'merge') },
					{ label: t('cmd.unmergeCells'), onSelect: () => mergeCells(store, 'unmerge') },
				]}
			/>
		</RibbonGroup>
	);
}
