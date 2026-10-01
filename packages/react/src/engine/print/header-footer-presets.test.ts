import { describe, expect, it } from 'vitest';

import { createTranslator } from '../../i18n/translator';
import { headerFooterPresets, type PresetLabels } from './header-footer-presets';

const labels: PresetLabels = { sheet: 'Sheet1', file: 'Book1.xlsx', date: '10/1/2026', page: '1' };

describe('headerFooterPresets', () => {
	it("writes Excel's built-in codes and captions in English", () => {
		const presets = headerFooterPresets(createTranslator(), labels);
		expect(presets.map((p) => p.code)).toEqual([
			'',
			'&CPage &P',
			'&CPage &P of &N',
			'&C&A',
			'&LConfidential&C&D&RPage &P',
			'&C&F',
			'&C&A&RPage &P',
			'&C&F&RPage &P',
			'&CPage &P&R&A',
			'&CPage &P&R&F',
		]);
		expect(presets.map((p) => p.label)).toEqual([
			'(none)',
			'Page 1',
			'Page 1 of ?',
			'Sheet1',
			'Confidential, 10/1/2026, Page 1',
			'Book1.xlsx',
			'Sheet1, Page 1',
			'Book1.xlsx, Page 1',
			'Page 1, Sheet1',
			'Page 1, Book1.xlsx',
		]);
	});

	it('translates the words written into the sheet and escapes ampersands', () => {
		const t = createTranslator({
			'dialog.pageSetup.presetNone': '(ohne)',
			'dialog.pageSetup.presetPage': 'Seite {page}',
			'dialog.pageSetup.presetPageOf': 'Seite {page} von {pages}',
			'dialog.pageSetup.presetConfidential': 'Vertraulich & intern',
		});
		const presets = headerFooterPresets(t, labels);
		expect(presets[0]).toEqual({ code: '', label: '(ohne)' });
		expect(presets[2]).toEqual({ code: '&CSeite &P von &N', label: 'Seite 1 von ?' });
		expect(presets[4]).toEqual({
			code: '&LVertraulich && intern&C&D&RSeite &P',
			label: 'Vertraulich & intern, 10/1/2026, Seite 1',
		});
	});
});
