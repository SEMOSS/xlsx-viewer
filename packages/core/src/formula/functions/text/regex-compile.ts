/**
 * Excel's REGEX functions use PCRE2. Most PCRE syntax is valid JavaScript;
 * this module translates the common differences (leading inline flags,
 * `\A` / `\z` / `\Z`, `\h`, POSIX classes, `(?P<name>...)` groups) and
 * compiles the result. Unsupported constructs make the pattern invalid.
 */

const POSIX_CLASSES: Record<string, string> = {
	alpha: 'a-zA-Z',
	digit: '0-9',
	alnum: 'a-zA-Z0-9',
	upper: 'A-Z',
	lower: 'a-z',
	space: '\\s',
	blank: ' \\t',
	punct: '!-\\/:-@\\[-`{-~',
	xdigit: '0-9A-Fa-f',
	word: '\\w',
	cntrl: '\\x00-\\x1f\\x7f',
	print: '\\x20-\\x7e',
	graph: '\\x21-\\x7e',
};

/** Translate PCRE-only syntax to JavaScript, or null when it cannot be expressed. */
export function translatePattern(src: string): string | null {
	let out = '';
	let inClass = false;
	for (let i = 0; i < src.length; i++) {
		const ch = src[i];
		if (ch === '\\') {
			const next = src[i + 1];
			if (next === undefined) {
				return null;
			}
			i++;
			if (next === 'h') {
				out += inClass ? ' \\t' : '[ \\t]';
			} else if (!inClass && next === 'A') {
				out += '(?<![\\s\\S])';
			} else if (!inClass && next === 'z') {
				out += '(?![\\s\\S])';
			} else if (!inClass && next === 'Z') {
				out += '(?=\\n?(?![\\s\\S]))';
			} else {
				out += ch + next;
			}
			continue;
		}
		if (inClass) {
			if (ch === '[' && src[i + 1] === ':') {
				const end = src.indexOf(':]', i + 2);
				const cls = end > 0 ? POSIX_CLASSES[src.slice(i + 2, end)] : undefined;
				if (cls === undefined) {
					return null;
				}
				out += cls;
				i = end + 1;
				continue;
			}
			if (ch === ']') {
				inClass = false;
			}
			out += ch;
			continue;
		}
		if (ch === '[') {
			inClass = true;
			out += ch;
			if (src[i + 1] === '^') {
				out += '^';
				i++;
			}
			// A `]` right after the opening bracket is a literal in PCRE.
			if (src[i + 1] === ']') {
				out += '\\]';
				i++;
			}
			continue;
		}
		if (src.startsWith('(?P<', i)) {
			out += '(?<';
			i += 3;
			continue;
		}
		if (src.startsWith('(?P=', i)) {
			const end = src.indexOf(')', i);
			if (end < 0) {
				return null;
			}
			out += `\\k<${src.slice(i + 4, end)}>`;
			i = end;
			continue;
		}
		out += ch;
	}
	return inClass ? null : out;
}

const cache = new Map<string, RegExp | null>();

/** Compile a PCRE-style pattern, or null when it is invalid. */
export function compilePattern(
	pattern: string,
	ignoreCase: boolean,
	global: boolean,
): RegExp | null {
	const key = `${ignoreCase ? 'i' : ''}${global ? 'g' : ''}:${pattern}`;
	if (cache.has(key)) {
		const re = cache.get(key)!;
		if (re) {
			re.lastIndex = 0;
		}
		return re;
	}
	let flags = ignoreCase ? 'i' : '';
	let src = pattern;
	const inline = /^\(\?([imsx]+)\)/.exec(src);
	if (inline) {
		if (inline[1].includes('x')) {
			cache.set(key, null);
			return null;
		}
		for (const f of inline[1]) {
			if (!flags.includes(f)) {
				flags += f;
			}
		}
		src = src.slice(inline[0].length);
	}
	if (global) {
		flags += 'g';
	}
	const translated = translatePattern(src);
	let re: RegExp | null = null;
	if (translated !== null) {
		for (const extra of ['u', '']) {
			try {
				re = new RegExp(translated, flags + extra);
				break;
			} catch {
				// Retry without the unicode flag (it rejects some legacy escapes).
			}
		}
	}
	if (cache.size > 200) {
		cache.clear();
	}
	cache.set(key, re);
	return re;
}

/**
 * Expand a replacement template for one match: `$n` / `${n}` insert a
 * capturing group, `${name}` a named group, `$0` / `$&` the whole match and
 * `$$` a literal dollar sign. Groups that did not participate insert "".
 */
export function expandReplacement(template: string, match: RegExpMatchArray): string {
	return template.replace(/\$(?:\$|&|(\d+)|\{(\w+)\})/g, (token, num?: string, braced?: string) => {
		if (token === '$$') {
			return '$';
		}
		if (token === '$&') {
			return match[0];
		}
		const key = num ?? braced ?? '';
		if (/^\d+$/.test(key)) {
			const index = Number(key);
			return index < match.length ? (match[index] ?? '') : '';
		}
		return match.groups?.[key] ?? '';
	});
}
