import { cleanup } from '@testing-library/react';
import { afterEach } from 'vitest';

// jsdom has no canvas: return null quietly so text measurement uses its estimate.
if (typeof HTMLCanvasElement !== 'undefined') {
	HTMLCanvasElement.prototype.getContext = (() =>
		null) as typeof HTMLCanvasElement.prototype.getContext;
}

afterEach(() => {
	cleanup();
});
