import { memo } from 'react';

import type { TableThumbnail } from '../../../engine/render/table-thumbnail';

/** SVG miniature of a table style (fills, text dashes and borders). */
export const TableThumbnailView = memo(function TableThumbnailView({
	thumb,
}: {
	thumb: TableThumbnail;
}) {
	return (
		<svg
			width={thumb.width + 2}
			height={thumb.height + 2}
			viewBox={`-1 -1 ${thumb.width + 2} ${thumb.height + 2}`}
			aria-hidden
		>
			<rect x={0} y={0} width={thumb.width} height={thumb.height} fill='#ffffff' />
			{thumb.cells.map((c, i) =>
				c.fill ? (
					<rect key={`f${i}`} x={c.x} y={c.y} width={c.width} height={c.height} fill={c.fill} />
				) : null,
			)}
			{thumb.cells.map((c, i) => (
				<line
					key={`t${i}`}
					x1={c.x + 2.5}
					x2={c.x + c.width - 3}
					y1={c.y + c.height / 2}
					y2={c.y + c.height / 2}
					stroke={c.text}
					strokeWidth={c.bold ? 1.6 : 1}
					opacity={0.85}
				/>
			))}
			{thumb.lines.map((l, i) => (
				<line
					key={`b${i}`}
					x1={l.x1}
					y1={l.y1}
					x2={l.x2}
					y2={l.y2}
					stroke={l.color}
					strokeWidth={l.width}
				/>
			))}
		</svg>
	);
});
