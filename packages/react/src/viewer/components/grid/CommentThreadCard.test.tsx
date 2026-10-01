import { render, screen } from '@testing-library/react';
import { createWorkbook } from 'xlsx-viewer-core';
import { describe, expect, it } from 'vitest';

import { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import { ViewerStoreContext } from '../../state/use-store';
import { ViewerStore } from '../../state/viewer-store';
import { CommentThreadCard } from './CommentThreadCard';

describe('CommentThreadCard', () => {
	it('shows the thread with authors and replies in order', () => {
		const wb = createWorkbook();
		wb.persons.push({ id: 'p1', displayName: 'Ana Silva' }, { id: 'p2', displayName: 'Ben Ito' });
		const sheet = wb.sheets[0];
		sheet.threadedComments.push(
			{
				id: 'c1',
				ref: { row: 1, col: 1 },
				personId: 'p1',
				text: 'Is this final?',
				created: '2026-03-01T10:00:00Z',
			},
			{
				id: 'c3',
				ref: { row: 1, col: 1 },
				personId: 'p1',
				text: 'Thanks!',
				created: '2026-03-02T09:00:00Z',
				parentId: 'c1',
			},
			{
				id: 'c2',
				ref: { row: 1, col: 1 },
				personId: 'p2',
				text: 'Yes, approved.',
				created: '2026-03-01T12:00:00Z',
				parentId: 'c1',
			},
		);
		const store = new ViewerStore(wb);
		render(
			<ViewerStoreContext.Provider value={store}>
				<CommentThreadCard
					root={sheet.threadedComments[0]}
					geometry={new SheetGeometry(sheet, { zoom: 1 })}
					rows={{ start: 0, end: 20 }}
					cols={{ start: 0, end: 10 }}
				/>
			</ViewerStoreContext.Provider>,
		);
		const card = screen.getByRole('dialog', { name: 'Comment thread' });
		expect(card.textContent).toMatch(
			/Ana Silva.*Is this final\?.*Ben Ito.*Yes, approved\..*Thanks!/s,
		);
		expect(screen.getAllByText('AS')).toHaveLength(2);
	});
});
