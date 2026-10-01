import type { DrawingObject, Workbook } from 'xlsx-viewer-core';
import { Component, memo, type ReactNode } from 'react';

import { useT } from '../../../i18n';
import type { Rect } from '../../../engine/geometry/sheet-geometry';
import { ChartView } from '../charts';
import { DRAWING_KIND_KEYS } from './drawing-labels';
import { isDisplayableImage, mediaUrl } from './media-url';
import { ShapeView } from './ShapeView';

/** Keeps one broken drawing from taking down the grid. */
class DrawingBoundary extends Component<
	{ label: string; children: ReactNode },
	{ failed: boolean }
> {
	override state = { failed: false };

	static getDerivedStateFromError(): { failed: boolean } {
		return { failed: true };
	}

	override render() {
		if (this.state.failed) {
			return <Placeholder label={this.props.label} />;
		}
		return this.props.children;
	}
}

function Placeholder({ label }: { label: string }) {
	return (
		<div className='flex h-full w-full items-center justify-center border border-dashed border-[#b8b8b8] bg-[#fafafa] p-1 text-center text-[11px] text-[#7a7a7a]'>
			{label}
		</div>
	);
}

interface Props {
	object: DrawingObject;
	rect: Rect;
	workbook: Workbook;
	sheetIndex: number;
	zoom: number;
	revision: number;
}

function Content({ object, rect, workbook, sheetIndex, zoom, revision }: Props): ReactNode {
	const t = useT();
	switch (object.kind) {
		case 'picture': {
			const media = workbook.media.get(object.mediaPath);
			const url =
				media && isDisplayableImage(media.contentType)
					? mediaUrl(workbook, object.mediaPath)
					: undefined;
			if (!url) {
				const label = media
					? t('drawing.pictureType', { type: media.contentType.replace('image/', '') })
					: t('drawing.picture');
				return <Placeholder label={label} />;
			}
			const c = object.crop;
			if (c && (c.l || c.t || c.r || c.b)) {
				const fw = 1 - ((c.l ?? 0) + (c.r ?? 0)) / 100000;
				const fh = 1 - ((c.t ?? 0) + (c.b ?? 0)) / 100000;
				const w = rect.width / Math.max(0.01, fw);
				const h = rect.height / Math.max(0.01, fh);
				return (
					<div style={{ position: 'absolute', inset: 0, overflow: 'hidden' }}>
						<img
							src={url}
							alt={object.description ?? object.name ?? ''}
							draggable={false}
							style={{
								position: 'absolute',
								width: w,
								height: h,
								left: (-w * (c.l ?? 0)) / 100000,
								top: (-h * (c.t ?? 0)) / 100000,
								maxWidth: 'none',
							}}
						/>
					</div>
				);
			}
			return (
				<img
					src={url}
					alt={object.description ?? object.name ?? ''}
					draggable={false}
					style={{ display: 'block', width: '100%', height: '100%', maxWidth: 'none' }}
				/>
			);
		}
		case 'chart':
			return (
				<ChartView
					chart={object.chart}
					workbook={workbook}
					sheetIndex={sheetIndex}
					width={Math.max(1, rect.width)}
					height={Math.max(1, rect.height)}
					revision={revision}
				/>
			);
		case 'shape':
			return (
				<ShapeView
					shape={object}
					width={rect.width}
					height={rect.height}
					theme={workbook.theme}
					zoom={zoom}
				/>
			);
		case 'group': {
			const off = object.childOffset ?? { x: 0, y: 0 };
			const ext = object.childExtent;
			return (
				<>
					{object.children.map((child) => {
						if (child.anchor.type !== 'absolute' || !ext || ext.cx === 0 || ext.cy === 0) {
							return null;
						}
						const sx = rect.width / ext.cx;
						const sy = rect.height / ext.cy;
						const childRect = {
							x: (child.anchor.pos.x - off.x) * sx,
							y: (child.anchor.pos.y - off.y) * sy,
							width: child.anchor.ext.cx * sx,
							height: child.anchor.ext.cy * sy,
						};
						return (
							<div
								key={child.id}
								style={{
									position: 'absolute',
									left: childRect.x,
									top: childRect.y,
									width: childRect.width,
									height: childRect.height,
								}}
							>
								<Content
									object={child}
									rect={childRect}
									workbook={workbook}
									sheetIndex={sheetIndex}
									zoom={zoom}
									revision={revision}
								/>
							</div>
						);
					})}
				</>
			);
		}
		default:
			return <Placeholder label={object.label ?? object.name ?? t('drawing.object')} />;
	}
}

/** One floating object (picture, chart, shape, group) at its anchor rectangle. */
export const DrawingObjectView = memo(function DrawingObjectView(props: Props) {
	const t = useT();
	const { object, rect } = props;
	const transforms: string[] = [];
	if (object.rotation) {
		transforms.push(`rotate(${object.rotation}deg)`);
	}
	if (object.flipH) {
		transforms.push('scaleX(-1)');
	}
	if (object.flipV) {
		transforms.push('scaleY(-1)');
	}
	return (
		<div
			className='xv-object-content'
			style={{
				position: 'absolute',
				inset: 0,
				transform: transforms.length ? transforms.join(' ') : undefined,
			}}
		>
			<DrawingBoundary label={object.name ?? t(DRAWING_KIND_KEYS[object.kind])}>
				<Content {...props} />
			</DrawingBoundary>
		</div>
	);
});
