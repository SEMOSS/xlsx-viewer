/** The twelve theme colors, as `RRGGBB` hex strings (no `#`). */
export interface ThemeColorScheme {
	name: string;
	dk1: string;
	lt1: string;
	dk2: string;
	lt2: string;
	accent1: string;
	accent2: string;
	accent3: string;
	accent4: string;
	accent5: string;
	accent6: string;
	hlink: string;
	folHlink: string;
}

export interface ThemeFontScheme {
	name: string;
	/** Headings font (latin). */
	majorLatin: string;
	/** Body font (latin). */
	minorLatin: string;
	majorEastAsian?: string;
	minorEastAsian?: string;
	majorComplexScript?: string;
	minorComplexScript?: string;
}

export interface Theme {
	name: string;
	colors: ThemeColorScheme;
	fonts: ThemeFontScheme;
	/** Original theme part XML, re-emitted unchanged on save. */
	rawXml?: string;
}
