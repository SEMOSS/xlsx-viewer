import { describe, expect, it } from 'vitest';

import { hyperlinkDescription, resolveHyperlink } from './hyperlinks';

describe('hyperlinks', () => {
	it('resolves safe URLs, locations and unsupported targets', () => {
		expect(
			resolveHyperlink({
				ref: { top: 0, left: 0, bottom: 0, right: 0 },
				target: 'https://x.example/a',
				location: 'top',
			}),
		).toEqual({
			kind: 'url',
			url: 'https://x.example/a#top',
		});
		expect(
			resolveHyperlink({
				ref: { top: 0, left: 0, bottom: 0, right: 0 },
				target: 'mailto:a@b.example',
			})?.kind,
		).toBe('url');
		expect(
			resolveHyperlink({
				ref: { top: 0, left: 0, bottom: 0, right: 0 },
				target: 'javascript:alert(1)',
			})?.kind,
		).toBe('unsupported');
		expect(
			resolveHyperlink({ ref: { top: 0, left: 0, bottom: 0, right: 0 }, target: 'Budget.xlsx' })
				?.kind,
		).toBe('unsupported');
		expect(
			resolveHyperlink({
				ref: { top: 0, left: 0, bottom: 0, right: 0 },
				location: "'My Sheet'!B2",
			}),
		).toEqual({
			kind: 'location',
			location: "'My Sheet'!B2",
		});
		expect(resolveHyperlink({ ref: { top: 0, left: 0, bottom: 0, right: 0 } })).toBeNull();
	});

	it('describes a link for its tooltip', () => {
		const ref = { top: 0, left: 0, bottom: 0, right: 0 };
		expect(hyperlinkDescription({ ref, target: 'https://x.example', tooltip: 'Docs' })).toBe(
			'Docs',
		);
		expect(hyperlinkDescription({ ref, target: 'https://x.example' })).toBe('https://x.example');
		expect(hyperlinkDescription({ ref, location: 'Data!A1' })).toBe('Data!A1');
	});
});
