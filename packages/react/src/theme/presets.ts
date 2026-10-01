import type { ViewerGridColors, ViewerTheme, ViewerThemeColors } from './types';

/**
 * Built-in themes. "Jade" uses Excel's green as the accent. The sheet canvas
 * is white in every preset (cells carry document colors), matching Excel.
 */

export const jadeLightColors: ViewerThemeColors = {
	background: '#f3f3f1',
	foreground: '#1b1d1c',
	card: '#ffffff',
	cardForeground: '#1b1d1c',
	popover: '#ffffff',
	popoverForeground: '#1b1d1c',
	primary: '#107c41',
	primaryForeground: '#ffffff',
	secondary: '#e9ebe8',
	secondaryForeground: '#1b1d1c',
	muted: '#ecedeb',
	mutedForeground: '#5b625e',
	accent: 'rgba(16, 124, 65, 0.09)',
	accentForeground: '#1b1d1c',
	destructive: '#c42b1c',
	destructiveForeground: '#ffffff',
	border: '#dcdedb',
	input: '#cfd2ce',
	ring: '#107c41',
};

export const jadeDarkColors: ViewerThemeColors = {
	background: '#111413',
	foreground: '#ecefed',
	card: '#181c1a',
	cardForeground: '#ecefed',
	popover: '#1c211f',
	popoverForeground: '#ecefed',
	primary: '#3fb27a',
	primaryForeground: '#06150d',
	secondary: '#222826',
	secondaryForeground: '#ecefed',
	muted: '#222826',
	mutedForeground: '#9ba6a0',
	accent: 'rgba(63, 178, 122, 0.12)',
	accentForeground: '#ecefed',
	destructive: '#f05a4f',
	destructiveForeground: '#ffffff',
	border: '#2c3330',
	input: '#343c38',
	ring: '#3fb27a',
};

export const lightGridColors: ViewerGridColors = {
	canvas: '#ffffff',
	text: '#000000',
	gridline: '#e1e1e1',
	headerBackground: '#f5f5f5',
	headerForeground: '#444444',
	headerBorder: '#d4d4d4',
	headerSelectedBackground: '#d3e7dc',
	headerSelectedForeground: '#0b5a2e',
	selectionBorder: '#107c41',
	selectionFill: 'rgba(16, 124, 65, 0.09)',
};

export const darkGridColors: ViewerGridColors = {
	canvas: '#ffffff',
	text: '#000000',
	gridline: '#e1e1e1',
	headerBackground: '#1c211f',
	headerForeground: '#b9c2bd',
	headerBorder: '#323a36',
	headerSelectedBackground: '#21402f',
	headerSelectedForeground: '#9be3bd',
	selectionBorder: '#1e9e5a',
	selectionFill: 'rgba(16, 124, 65, 0.11)',
};

export const DEFAULT_RADIUS = '0.375rem';
export const DEFAULT_UI_FONT =
	"'Segoe UI Variable Text', 'Segoe UI', system-ui, -apple-system, 'Helvetica Neue', Arial, sans-serif";

export const jadeLightTheme: ViewerTheme = {
	mode: 'light',
	colors: jadeLightColors,
	grid: lightGridColors,
	radius: DEFAULT_RADIUS,
};

export const jadeDarkTheme: ViewerTheme = {
	mode: 'dark',
	colors: jadeDarkColors,
	grid: darkGridColors,
	radius: DEFAULT_RADIUS,
};

export interface ThemeCatalogEntry {
	key: string;
	label: string;
	theme: ViewerTheme;
}

/** Themes offered by the built-in theme switcher. */
export const BUILTIN_THEMES: readonly ThemeCatalogEntry[] = [
	{ key: 'jadeLight', label: 'Light', theme: jadeLightTheme },
	{ key: 'jadeDark', label: 'Dark', theme: jadeDarkTheme },
];
