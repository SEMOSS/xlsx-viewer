import { describe, expect, it } from 'vitest';

import { createTranslator } from '../../../i18n';
import { DRAWING_KIND_KEYS } from './drawing-labels';

describe('DRAWING_KIND_KEYS', () => {
	it('names each kind of object in English as the kind itself', () => {
		const t = createTranslator();
		for (const [kind, key] of Object.entries(DRAWING_KIND_KEYS)) {
			expect(t(key)).toBe(kind);
		}
	});
});
