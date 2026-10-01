import type { ThreadedComment } from 'xlsx-viewer-core';

import { useLocale, useT } from '../../../i18n';
import type { SheetGeometry } from '../../../engine/geometry/sheet-geometry';
import type { Span } from '../../../engine/layout/types';
import { useViewerStore } from '../../state/use-store';

const AVATAR_COLORS = [
	'#7160e8',
	'#0f7b6c',
	'#c43e1c',
	'#2b5797',
	'#8764b8',
	'#a4373a',
	'#0b6a0b',
	'#986f0b',
];

function initials(name: string): string {
	const parts = name.trim().split(/\s+/).filter(Boolean);
	return (
		(parts[0]?.[0] ?? '?') + (parts.length > 1 ? parts[parts.length - 1][0] : '')
	).toUpperCase();
}

function colorFor(name: string): string {
	let h = 0;
	for (const ch of name) {
		h = (h * 31 + ch.charCodeAt(0)) >>> 0;
	}
	return AVATAR_COLORS[h % AVATAR_COLORS.length];
}

function when(iso: string | undefined, locale: string): string {
	if (!iso) {
		return '';
	}
	const d = new Date(iso);
	return Number.isNaN(d.getTime())
		? ''
		: d.toLocaleString(locale, { dateStyle: 'medium', timeStyle: 'short' });
}

/** A threaded comment (Excel 365 comments) shown as a card while its cell is hovered. */
export function CommentThreadCard({
	root,
	geometry: g,
	rows,
	cols,
}: {
	root: ThreadedComment;
	geometry: SheetGeometry;
	rows: Span;
	cols: Span;
}) {
	const t = useT();
	const locale = useLocale();
	const store = useViewerStore();
	const { row, col } = root.ref;
	if (row < rows.start || row > rows.end || col < cols.start || col > cols.end) {
		return null;
	}
	const sheet = store.activeSheet;
	const replies = sheet.threadedComments
		.filter((c) => c.parentId === root.id)
		.sort((a, b) => (a.created ?? '').localeCompare(b.created ?? ''));
	const persons = new Map(store.workbook.persons.map((p) => [p.id, p.displayName]));
	const thread = [root, ...replies];
	return (
		<div
			className='xv-comment-card'
			role='dialog'
			aria-label={t('comment.thread')}
			style={{ left: g.colLeft(col) + g.colWidth(col) + 10, top: Math.max(0, g.rowTop(row) - 4) }}
		>
			{root.done && <div className='xv-comment-resolved'>{t('comment.resolved')}</div>}
			{thread.map((c) => {
				const name = persons.get(c.personId) ?? t('comment.unknownAuthor');
				return (
					<div key={c.id} className='xv-comment-entry'>
						<span className='xv-comment-avatar' style={{ background: colorFor(name) }} aria-hidden>
							{initials(name)}
						</span>
						<div className='min-w-0'>
							<div className='xv-comment-meta'>
								<strong>{name}</strong>
								<span>{when(c.created, locale)}</span>
							</div>
							<div className='xv-comment-text'>{c.text}</div>
						</div>
					</div>
				);
			})}
		</div>
	);
}
