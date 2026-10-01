import { memo, type CSSProperties } from 'react';
import { ICON_SETS, type IconSetName } from 'xlsx-viewer-core';

import { useT } from '../../../i18n';
import { ICON_SET_KEYS } from './cf-labels';
import { renderIconShape } from './icon-shapes';

export interface IconSetIconProps {
	iconSet: IconSetName;
	/** Icon index within the set (0 = icon for the lowest values), as in `IconResult`. */
	index: number;
	/** Rendered size in CSS pixels. The icon is drawn on a 16-unit grid, so 16 is pixel-exact. */
	size?: number;
	className?: string;
	style?: CSSProperties;
}

function IconSetIconImpl({ iconSet, index, size = 16, className, style }: IconSetIconProps) {
	const t = useT();
	const set = ICON_SETS[iconSet];
	const icon = set?.icons[index];
	if (!icon) {
		return null;
	}
	return (
		<svg
			className={className}
			width={size}
			height={size}
			viewBox='0 0 16 16'
			role='img'
			aria-label={t('grid.iconSetIcon', {
				name: t(ICON_SET_KEYS[iconSet]),
				index: index + 1,
				count: set.count,
			})}
			focusable='false'
			data-icon-set={iconSet}
			data-icon-index={index}
			data-shape={icon.shape}
			style={{ display: 'block', flexShrink: 0, ...style }}
		>
			{renderIconShape(icon)}
		</svg>
	);
}

/** An Excel icon-set icon (conditional formatting) as an inline SVG. */
export const IconSetIcon = memo(IconSetIconImpl);
IconSetIcon.displayName = 'IconSetIcon';
