import {
	DEFAULT_RADIUS,
	DEFAULT_UI_FONT,
	darkGridColors,
	jadeDarkColors,
	jadeLightColors,
	lightGridColors,
} from './presets';
import type { ViewerGridColors, ViewerTheme, ViewerThemeColors } from './types';

const COLOR_VARS: Record<keyof ViewerThemeColors, string> = {
	background: '--xlsx-background',
	foreground: '--xlsx-foreground',
	card: '--xlsx-card',
	cardForeground: '--xlsx-card-foreground',
	popover: '--xlsx-popover',
	popoverForeground: '--xlsx-popover-foreground',
	primary: '--xlsx-primary',
	primaryForeground: '--xlsx-primary-foreground',
	secondary: '--xlsx-secondary',
	secondaryForeground: '--xlsx-secondary-foreground',
	muted: '--xlsx-muted',
	mutedForeground: '--xlsx-muted-foreground',
	accent: '--xlsx-accent',
	accentForeground: '--xlsx-accent-foreground',
	destructive: '--xlsx-destructive',
	destructiveForeground: '--xlsx-destructive-foreground',
	border: '--xlsx-border',
	input: '--xlsx-input',
	ring: '--xlsx-ring',
};

const GRID_VARS: Record<keyof ViewerGridColors, string> = {
	canvas: '--xlsx-grid-canvas',
	text: '--xlsx-grid-text',
	gridline: '--xlsx-gridline',
	headerBackground: '--xlsx-header-bg',
	headerForeground: '--xlsx-header-fg',
	headerBorder: '--xlsx-header-border',
	headerSelectedBackground: '--xlsx-header-selected-bg',
	headerSelectedForeground: '--xlsx-header-selected-fg',
	selectionBorder: '--xlsx-selection-border',
	selectionFill: '--xlsx-selection-fill',
};

export interface ResolvedViewerTheme {
	mode: 'light' | 'dark';
	colors: ViewerThemeColors;
	grid: ViewerGridColors;
	radius: string;
	fontFamily: string;
}

/** Resolve a (partial) theme against the defaults for its mode. */
export function resolveTheme(theme: ViewerTheme | undefined): ResolvedViewerTheme {
	const mode = theme?.mode ?? 'light';
	return {
		mode,
		colors: { ...(mode === 'dark' ? jadeDarkColors : jadeLightColors), ...(theme?.colors ?? {}) },
		grid: { ...(mode === 'dark' ? darkGridColors : lightGridColors), ...(theme?.grid ?? {}) },
		radius: theme?.radius ?? DEFAULT_RADIUS,
		fontFamily: theme?.fontFamily ?? DEFAULT_UI_FONT,
	};
}

/** CSS custom properties for a theme, ready for a `style` attribute or `:root`. */
export function themeToCssVars(theme: ViewerTheme | undefined): Record<string, string> {
	const t = resolveTheme(theme);
	const vars: Record<string, string> = {};
	for (const [key, cssVar] of Object.entries(COLOR_VARS) as Array<
		[keyof ViewerThemeColors, string]
	>) {
		vars[cssVar] = t.colors[key];
	}
	for (const [key, cssVar] of Object.entries(GRID_VARS) as Array<
		[keyof ViewerGridColors, string]
	>) {
		vars[cssVar] = t.grid[key];
	}
	vars['--xlsx-radius'] = t.radius;
	vars['--xlsx-font'] = t.fontFamily;
	return vars;
}
