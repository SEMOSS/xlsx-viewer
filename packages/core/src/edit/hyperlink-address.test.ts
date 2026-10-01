import { describe, expect, it } from 'vitest';

import {
	buildMailtoAddress,
	formatHyperlinkLocation,
	normalizeHyperlinkAddress,
	parseHyperlinkLocation,
	parseMailtoAddress,
} from './index';

describe('normalizeHyperlinkAddress', () => {
	it('completes web, ftp and e-mail addresses like Excel', () => {
		expect(normalizeHyperlinkAddress('www.example.com')).toBe('http://www.example.com');
		expect(normalizeHyperlinkAddress('  WWW.Example.com/a?b=1 ')).toBe(
			'http://WWW.Example.com/a?b=1',
		);
		expect(normalizeHyperlinkAddress('www.example.com:8080')).toBe('http://www.example.com:8080');
		expect(normalizeHyperlinkAddress('ftp.example.com')).toBe('ftp://ftp.example.com');
		expect(normalizeHyperlinkAddress('user@example.com')).toBe('mailto:user@example.com');
		expect(normalizeHyperlinkAddress('first.last@mail.example.org')).toBe(
			'mailto:first.last@mail.example.org',
		);
	});

	it('keeps addresses with a scheme and file paths', () => {
		for (const address of [
			'https://example.com/x',
			'http://localhost:3000',
			'mailto:a@b.com?subject=Hi',
			'file:///C:/Docs/a.xlsx',
			'C:\\Docs\\a.xlsx',
			'\\\\server\\share\\a.xlsx',
			'report.xlsx',
			'../data/b.xlsx',
			'/var/data.csv',
			'#Sheet2!A1',
		]) {
			expect(normalizeHyperlinkAddress(address)).toBe(address);
		}
		expect(normalizeHyperlinkAddress('   ')).toBe('');
	});
});

describe('mailto addresses', () => {
	it('builds mailto links with an encoded subject', () => {
		expect(buildMailtoAddress('a@b.com')).toBe('mailto:a@b.com');
		expect(buildMailtoAddress(' mailto:a@b.com ', 'Hello there')).toBe(
			'mailto:a@b.com?subject=Hello%20there',
		);
		expect(buildMailtoAddress('a@b.com', 'Q&A?')).toBe('mailto:a@b.com?subject=Q%26A%3F');
		expect(buildMailtoAddress('a@b.com', '   ')).toBe('mailto:a@b.com');
	});

	it('parses mailto links back', () => {
		expect(parseMailtoAddress('mailto:a@b.com?subject=Hello%20there')).toEqual({
			email: 'a@b.com',
			subject: 'Hello there',
		});
		expect(parseMailtoAddress('MAILTO:a@b.com?cc=c@d.com&Subject=Hi+you')).toEqual({
			email: 'a@b.com',
			subject: 'Hi you',
		});
		expect(parseMailtoAddress('mailto:a@b.com?subject=100%')).toEqual({
			email: 'a@b.com',
			subject: '100%',
		});
		expect(parseMailtoAddress('https://example.com')).toBeNull();
	});
});

describe('hyperlink locations', () => {
	it('splits sheet and reference', () => {
		expect(parseHyperlinkLocation('Sheet2!B4')).toEqual({ sheet: 'Sheet2', ref: 'B4' });
		expect(parseHyperlinkLocation("'My Sheet'!A1:C3")).toEqual({ sheet: 'My Sheet', ref: 'A1:C3' });
		expect(parseHyperlinkLocation("'Bob''s'!A1")).toEqual({ sheet: "Bob's", ref: 'A1' });
		expect(parseHyperlinkLocation('#Sheet1!A1')).toEqual({ sheet: 'Sheet1', ref: 'A1' });
		expect(parseHyperlinkLocation('Rate')).toEqual({ ref: 'Rate' });
		expect(parseHyperlinkLocation("'Broken")).toEqual({ ref: "'Broken" });
	});

	it('formats locations with quoted sheet names', () => {
		expect(formatHyperlinkLocation('Sheet2', 'A1')).toBe('Sheet2!A1');
		expect(formatHyperlinkLocation('My Sheet', 'B2:C3')).toBe("'My Sheet'!B2:C3");
		expect(formatHyperlinkLocation(undefined, 'Rate')).toBe('Rate');
	});
});
