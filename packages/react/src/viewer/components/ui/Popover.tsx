import {
	createContext,
	useCallback,
	useContext,
	useEffect,
	useLayoutEffect,
	useRef,
	useState,
	type ReactNode,
	type RefObject,
} from 'react';
import { createPortal } from 'react-dom';

import { isGridKeyTarget } from '../../focus';

/** Element inside the viewer root where floating UI is portalled (so theme variables apply). */
export const PortalContext = createContext<HTMLElement | null>(null);

export function usePortalHost(): HTMLElement | null {
	return useContext(PortalContext);
}

export type Placement = 'bottom-start' | 'bottom-end' | 'right-start' | 'top-start' | 'point';

interface PopoverProps {
	open: boolean;
	onClose: () => void;
	/** Element the popover is attached to (ignored for `point`). */
	anchorRef?: RefObject<HTMLElement | null>;
	/** Viewport coordinates for `point` placement (context menus). */
	point?: { x: number; y: number };
	placement?: Placement;
	children: ReactNode;
	className?: string;
	/** Keep open when clicking inside these elements (e.g. the anchor). */
	ignoreRefs?: Array<RefObject<HTMLElement | null>>;
	ariaLabel?: string;
	role?: string;
}

/** A floating panel positioned next to an anchor, closed on outside click / Escape. */
export function Popover({
	open,
	onClose,
	anchorRef,
	point,
	placement = 'bottom-start',
	children,
	className,
	ignoreRefs,
	ariaLabel,
	role = 'dialog',
}: PopoverProps) {
	const host = usePortalHost();
	const panelRef = useRef<HTMLDivElement>(null);
	const [pos, setPos] = useState<{ left: number; top: number } | null>(null);

	const place = useCallback(() => {
		const panel = panelRef.current;
		if (!panel) {
			return;
		}
		const pw = panel.offsetWidth;
		const ph = panel.offsetHeight;
		const vw = window.innerWidth;
		const vh = window.innerHeight;
		let left = 0;
		let top = 0;
		if (placement === 'point' && point) {
			left = point.x;
			top = point.y;
		} else {
			const a = anchorRef?.current?.getBoundingClientRect();
			if (!a) {
				return;
			}
			switch (placement) {
				case 'bottom-end':
					left = a.right - pw;
					top = a.bottom + 2;
					break;
				case 'right-start':
					left = a.right - 2;
					top = a.top - 4;
					break;
				case 'top-start':
					left = a.left;
					top = a.top - ph - 2;
					break;
				default:
					left = a.left;
					top = a.bottom + 2;
			}
			if (top + ph > vh - 4 && placement !== 'top-start') {
				top = placement === 'right-start' ? Math.max(4, vh - ph - 4) : Math.max(4, a.top - ph - 2);
			}
			if (placement === 'right-start' && left + pw > vw - 4) {
				left = Math.max(4, a.left - pw + 2);
			}
		}
		if (left + pw > vw - 4) {
			left = Math.max(4, vw - pw - 4);
		}
		if (top + ph > vh - 4) {
			top = Math.max(4, vh - ph - 4);
		}
		setPos({ left: Math.max(4, left), top: Math.max(4, top) });
	}, [anchorRef, placement, point]);

	useLayoutEffect(() => {
		if (open) {
			place();
		} else {
			setPos(null);
		}
	}, [open, place]);

	// Closing hands focus back to where it was when the popover opened, if the
	// panel had it (it is removed with the panel).
	useEffect(() => {
		if (!open) {
			return;
		}
		const prev = document.activeElement;
		const panel = panelRef.current;
		return () => {
			const active = document.activeElement;
			const lost = !active || active === document.body || !!panel?.contains(active);
			if (lost && prev instanceof HTMLElement && prev.isConnected && !panel?.contains(prev)) {
				prev.focus({ preventScroll: true });
			}
		};
	}, [open]);

	// Stay inside the window when it is resized while open.
	useEffect(() => {
		if (!open) {
			return;
		}
		window.addEventListener('resize', place);
		return () => window.removeEventListener('resize', place);
	}, [open, place]);

	useEffect(() => {
		if (!open) {
			return;
		}
		const onDown = (e: PointerEvent) => {
			const t = e.target as Node;
			if (panelRef.current?.contains(t)) {
				return;
			}
			if (anchorRef?.current?.contains(t)) {
				return;
			}
			if (ignoreRefs?.some((r) => r.current?.contains(t))) {
				return;
			}
			// Nested popovers are portalled as siblings; keep the parent open while they are used.
			if (
				(t as Element).closest?.('[data-xv-popover]') &&
				(t as Element).closest('[data-xv-popover]') !== panelRef.current
			) {
				return;
			}
			onClose();
		};
		const onKey = (e: KeyboardEvent) => {
			if (e.key === 'Escape') {
				e.stopPropagation();
				onClose();
			} else if (
				isGridKeyTarget(e.target) &&
				!['Shift', 'Control', 'Alt', 'Meta'].includes(e.key)
			) {
				onClose(); // typing in the grid leaves the dropdown
			}
		};
		const onResize = () => place();
		document.addEventListener('pointerdown', onDown, true);
		document.addEventListener('keydown', onKey, true);
		window.addEventListener('resize', onResize);
		return () => {
			document.removeEventListener('pointerdown', onDown, true);
			document.removeEventListener('keydown', onKey, true);
			window.removeEventListener('resize', onResize);
		};
	}, [open, onClose, anchorRef, ignoreRefs, place]);

	if (!open) {
		return null;
	}
	const panel = (
		<div
			ref={panelRef}
			data-xv-popover
			role={role}
			aria-label={ariaLabel}
			className={`xv-popover fixed z-[1000] rounded-md border border-border bg-popover text-popover-foreground shadow-lg ${className ?? ''}`}
			style={{
				left: pos?.left ?? -9999,
				top: pos?.top ?? -9999,
				visibility: pos ? 'visible' : 'hidden',
			}}
		>
			{children}
		</div>
	);
	return host ? createPortal(panel, host) : panel;
}
