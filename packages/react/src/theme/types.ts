/** Semantic colors of the viewer chrome (ribbon, dialogs, headers). Any CSS color string. */
export interface ViewerThemeColors {
	background: string;
	foreground: string;
	card: string;
	cardForeground: string;
	popover: string;
	popoverForeground: string;
	primary: string;
	primaryForeground: string;
	secondary: string;
	secondaryForeground: string;
	muted: string;
	mutedForeground: string;
	accent: string;
	accentForeground: string;
	destructive: string;
	destructiveForeground: string;
	border: string;
	input: string;
	ring: string;
}

/** Colors of the spreadsheet surface itself. The sheet canvas stays paper-white like Excel by default. */
export interface ViewerGridColors {
	/** Sheet background (cells without a fill). */
	canvas: string;
	/** Default text color for automatic font colors. */
	text: string;
	gridline: string;
	headerBackground: string;
	headerForeground: string;
	headerBorder: string;
	/** Header of selected rows / columns. */
	headerSelectedBackground: string;
	headerSelectedForeground: string;
	selectionBorder: string;
	selectionFill: string;
}

export interface ViewerTheme {
	/** Light or dark chrome; drives `color-scheme` for native controls. */
	mode?: 'light' | 'dark';
	colors?: Partial<ViewerThemeColors>;
	grid?: Partial<ViewerGridColors>;
	/** Base corner radius (CSS length), default `0.375rem`. */
	radius?: string;
	/** UI font stack for the chrome. */
	fontFamily?: string;
}
