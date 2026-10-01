import { Languages } from 'lucide-react';
import type { CSSProperties } from 'react';
import { BUILT_IN_LOCALES, isSupportedLocale, type SupportedLocale } from 'xlsx-react-viewer';

interface LanguagePickerProps {
	value: SupportedLocale;
	onChange: (locale: SupportedLocale) => void;
	/** Extra class, e.g. `demo-locale-floating` while a workbook is open. */
	className?: string;
	style?: CSSProperties;
}

/** Select for the viewer's UI language, listing each built-in locale by its native name. */
export function LanguagePicker({ value, onChange, className, style }: LanguagePickerProps) {
	return (
		<label className={`demo-locale ${className ?? ''}`} style={style} title='Language'>
			<Languages size={14} aria-hidden />
			<select
				aria-label='Language'
				value={value}
				onChange={(e) => {
					if (isSupportedLocale(e.target.value)) {
						onChange(e.target.value);
					}
				}}
			>
				{BUILT_IN_LOCALES.map((locale) => (
					<option key={locale.code} value={locale.code} lang={locale.code}>
						{locale.nativeName}
					</option>
				))}
			</select>
		</label>
	);
}
