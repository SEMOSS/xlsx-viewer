import { useRef, type KeyboardEvent, type PointerEvent } from 'react';

import { useT } from '../../../../i18n';

/** Stored `textRotation` (0-90 up, 91-180 down, 255 stacked) to dial degrees (-90..90). */
export function rotationToDegrees(rotation: number): number {
	if (rotation <= 90) {
		return rotation;
	}
	return rotation <= 180 ? 90 - rotation : 0;
}

/** Dial degrees (-90..90) to the stored `textRotation`. */
export function degreesToRotation(degrees: number): number {
	const d = Math.max(-90, Math.min(90, Math.round(degrees)));
	return d >= 0 ? d : 90 - d;
}

const W = 112;
const H = 132;
const CX = 14;
const CY = H / 2;
const R = 54;

interface OrientationDialProps {
	/** Current angle in degrees, or null when stacked / mixed. */
	degrees: number | null;
	onChange: (degrees: number) => void;
	disabled?: boolean;
}

/**
 * Excel's orientation half-dial: drag or click to set the text angle,
 * arrow keys change it by one degree (Page Up / Down by 15).
 */
export function OrientationDial({ degrees, onChange, disabled }: OrientationDialProps) {
	const t = useT();
	const ref = useRef<SVGSVGElement>(null);
	const dragging = useRef(false);
	const angle = degrees ?? 0;
	const rad = (angle * Math.PI) / 180;
	const tipX = CX + R * Math.cos(rad);
	const tipY = CY - R * Math.sin(rad);

	const fromPointer = (e: PointerEvent<SVGSVGElement>) => {
		const box = ref.current?.getBoundingClientRect();
		if (!box || box.width === 0) {
			return;
		}
		const x = ((e.clientX - box.left) / box.width) * W - CX;
		const y = CY - ((e.clientY - box.top) / box.height) * H;
		const deg = (Math.atan2(y, Math.max(0, x)) * 180) / Math.PI;
		onChange(Math.max(-90, Math.min(90, Math.round(deg))));
	};

	const onKeyDown = (e: KeyboardEvent<SVGSVGElement>) => {
		const steps: Record<string, number> = {
			PageUp: 15,
			PageDown: -15,
			ArrowUp: 1,
			ArrowRight: 1,
			ArrowDown: -1,
			ArrowLeft: -1,
		};
		const step = steps[e.key] ?? 0;
		if (e.key === 'Home' || e.key === 'End') {
			e.preventDefault();
			onChange(e.key === 'Home' ? -90 : 90);
			return;
		}
		if (step) {
			e.preventDefault();
			onChange(Math.max(-90, Math.min(90, angle + step)));
		}
	};

	const marks = [];
	for (let a = -90; a <= 90; a += 15) {
		const r = (a * Math.PI) / 180;
		marks.push(
			<circle
				key={a}
				cx={CX + R * Math.cos(r)}
				cy={CY - R * Math.sin(r)}
				r={a % 45 === 0 ? 2 : 1.4}
				className='fill-muted-foreground'
			/>,
		);
	}

	return (
		<svg
			ref={ref}
			width={W}
			height={H}
			viewBox={`0 0 ${W} ${H}`}
			role='slider'
			tabIndex={disabled ? -1 : 0}
			aria-label={t('dialog.formatCells.textOrientation')}
			aria-valuemin={-90}
			aria-valuemax={90}
			aria-valuenow={angle}
			aria-valuetext={t('dialog.formatCells.degreesValue', { degrees: angle })}
			aria-disabled={disabled || undefined}
			className={`touch-none select-none rounded-sm outline-none focus-visible:ring-2 focus-visible:ring-ring ${disabled ? 'opacity-40' : ''}`}
			onKeyDown={disabled ? undefined : onKeyDown}
			onPointerDown={(e) => {
				if (disabled) {
					return;
				}
				dragging.current = true;
				e.currentTarget.setPointerCapture?.(e.pointerId);
				fromPointer(e);
			}}
			onPointerMove={(e) => dragging.current && !disabled && fromPointer(e)}
			onPointerUp={() => {
				dragging.current = false;
			}}
		>
			<path
				d={`M ${CX} ${CY - R - 6} A ${R + 6} ${R + 6} 0 0 1 ${CX} ${CY + R + 6}`}
				className='fill-none stroke-border'
			/>
			{marks}
			{degrees !== null && (
				<>
					<line
						x1={CX + 34 * Math.cos(rad)}
						y1={CY - 34 * Math.sin(rad)}
						x2={tipX}
						y2={tipY}
						className='stroke-foreground'
						strokeWidth={1}
					/>
					<text
						x={CX + 8}
						y={CY + 4}
						fontSize={11}
						className='fill-foreground'
						transform={`rotate(${-angle} ${CX} ${CY})`}
					>
						{t('dialog.formatCells.previewText')}
					</text>
					<rect
						x={tipX - 3.5}
						y={tipY - 3.5}
						width={7}
						height={7}
						transform={`rotate(45 ${tipX} ${tipY})`}
						className='fill-destructive'
					/>
				</>
			)}
		</svg>
	);
}
