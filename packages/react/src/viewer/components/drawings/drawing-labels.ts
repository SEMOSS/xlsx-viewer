import type { DrawingObject } from 'xlsx-viewer-core';

import type { TranslationKey } from '../../../i18n';

/** What a floating object is called when it has no name or alt text, by kind. */
export const DRAWING_KIND_KEYS: Readonly<Record<DrawingObject['kind'], TranslationKey>> = {
	picture: 'drawing.kind.picture',
	chart: 'drawing.kind.chart',
	shape: 'drawing.kind.shape',
	group: 'drawing.kind.group',
	unknown: 'drawing.kind.unknown',
};
