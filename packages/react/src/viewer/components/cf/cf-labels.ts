import type { IconSetName } from 'xlsx-viewer-core';

import type { Translate, TranslationKey } from '../../../i18n';

/**
 * Translation keys of the names the core gives its conditional formatting
 * presets (in English). The galleries look names up here and fall back to the
 * core's wording for a preset without a key.
 */
type PresetKeys = Readonly<Partial<Record<string, TranslationKey>>>;

/** `DATA_BAR_PRESETS` names. */
export const DATA_BAR_KEYS: PresetKeys = {
	'Blue Data Bar': 'cf.dataBar.blue',
	'Green Data Bar': 'cf.dataBar.green',
	'Red Data Bar': 'cf.dataBar.red',
	'Orange Data Bar': 'cf.dataBar.orange',
	'Light Blue Data Bar': 'cf.dataBar.lightBlue',
	'Purple Data Bar': 'cf.dataBar.purple',
};

/** `COLOR_SCALE_PRESETS` names. */
export const COLOR_SCALE_KEYS: PresetKeys = {
	'Green - Yellow - Red Color Scale': 'cf.colorScale.greenYellowRed',
	'Red - Yellow - Green Color Scale': 'cf.colorScale.redYellowGreen',
	'Green - White - Red Color Scale': 'cf.colorScale.greenWhiteRed',
	'Red - White - Green Color Scale': 'cf.colorScale.redWhiteGreen',
	'Blue - White - Red Color Scale': 'cf.colorScale.blueWhiteRed',
	'Red - White - Blue Color Scale': 'cf.colorScale.redWhiteBlue',
	'White - Red Color Scale': 'cf.colorScale.whiteRed',
	'Red - White Color Scale': 'cf.colorScale.redWhite',
	'Green - White Color Scale': 'cf.colorScale.greenWhite',
	'White - Green Color Scale': 'cf.colorScale.whiteGreen',
	'Green - Yellow Color Scale': 'cf.colorScale.greenYellow',
	'Yellow - Green Color Scale': 'cf.colorScale.yellowGreen',
};

/** `ICON_SET_GROUPS` headings. */
export const ICON_SET_GROUP_KEYS: PresetKeys = {
	Directional: 'cf.iconGroup.directional',
	Shapes: 'cf.iconGroup.shapes',
	Indicators: 'cf.iconGroup.indicators',
	Ratings: 'cf.iconGroup.ratings',
};

/** Gallery name of every icon set (`ICON_SETS[name].label`). */
export const ICON_SET_KEYS: Readonly<Record<IconSetName, TranslationKey>> = {
	'3Arrows': 'cf.iconSet.arrows3',
	'3ArrowsGray': 'cf.iconSet.arrows3Gray',
	'3Triangles': 'cf.iconSet.triangles3',
	'4Arrows': 'cf.iconSet.arrows4',
	'4ArrowsGray': 'cf.iconSet.arrows4Gray',
	'5Arrows': 'cf.iconSet.arrows5',
	'5ArrowsGray': 'cf.iconSet.arrows5Gray',
	'3TrafficLights1': 'cf.iconSet.trafficLights3Unrimmed',
	'3TrafficLights2': 'cf.iconSet.trafficLights3Rimmed',
	'3Signs': 'cf.iconSet.signs3',
	'4TrafficLights': 'cf.iconSet.trafficLights4',
	'4RedToBlack': 'cf.iconSet.redToBlack4',
	'3Symbols': 'cf.iconSet.symbols3Circled',
	'3Symbols2': 'cf.iconSet.symbols3Uncircled',
	'3Flags': 'cf.iconSet.flags3',
	'3Stars': 'cf.iconSet.stars3',
	'4Rating': 'cf.iconSet.ratings4',
	'5Quarters': 'cf.iconSet.quarters5',
	'5Rating': 'cf.iconSet.ratings5',
	'5Boxes': 'cf.iconSet.boxes5',
	NoIcons: 'cf.iconSet.none',
};

/** `name` (a core preset name) in the UI language, through `keys`. */
export function presetLabel(keys: PresetKeys, name: string, t: Translate): string {
	const key = keys[name];
	return key ? t(key) : name;
}
