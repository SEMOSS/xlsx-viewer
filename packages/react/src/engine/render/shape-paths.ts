/**
 * SVG paths for common DrawingML preset geometries, in a `w` x `h` box.
 * Adjust values (`adj`, in 1/100000 of the shorter side unless noted) use
 * Excel's defaults when absent. Unknown presets fall back to a rectangle.
 */
export function presetPath(
	preset: string,
	w: number,
	h: number,
	adjust: Record<string, number> = {},
): string {
	const ss = Math.min(w, h);
	const adj = (name: string, def: number) => (adjust[name] ?? def) / 100000;
	const poly = (pts: Array<[number, number]>) => `M${pts.map(([x, y]) => `${x} ${y}`).join('L')}Z`;
	switch (preset) {
		case 'rect':
		case 'flowChartProcess':
			return `M0 0H${w}V${h}H0Z`;
		case 'roundRect':
		case 'flowChartAlternateProcess': {
			const r = Math.min(ss * adj('adj', 16667), w / 2, h / 2);
			return `M${r} 0H${w - r}A${r} ${r} 0 0 1 ${w} ${r}V${h - r}A${r} ${r} 0 0 1 ${w - r} ${h}H${r}A${r} ${r} 0 0 1 0 ${h - r}V${r}A${r} ${r} 0 0 1 ${r} 0Z`;
		}
		case 'ellipse':
		case 'flowChartConnector':
			return `M0 ${h / 2}A${w / 2} ${h / 2} 0 1 0 ${w} ${h / 2}A${w / 2} ${h / 2} 0 1 0 0 ${h / 2}Z`;
		case 'triangle': {
			const x = w * adj('adj', 50000);
			return poly([
				[x, 0],
				[w, h],
				[0, h],
			]);
		}
		case 'rtTriangle':
			return poly([
				[0, 0],
				[w, h],
				[0, h],
			]);
		case 'diamond':
		case 'flowChartDecision':
			return poly([
				[w / 2, 0],
				[w, h / 2],
				[w / 2, h],
				[0, h / 2],
			]);
		case 'parallelogram':
		case 'flowChartInputOutput': {
			const o = ss * adj('adj', 25000);
			return poly([
				[o, 0],
				[w, 0],
				[w - o, h],
				[0, h],
			]);
		}
		case 'trapezoid': {
			const o = ss * adj('adj', 25000);
			return poly([
				[o, 0],
				[w - o, 0],
				[w, h],
				[0, h],
			]);
		}
		case 'pentagon':
		case 'homePlate': {
			const o = ss * adj('adj', 50000);
			return poly([
				[0, 0],
				[w - o, 0],
				[w, h / 2],
				[w - o, h],
				[0, h],
			]);
		}
		case 'chevron': {
			const o = ss * adj('adj', 50000);
			return poly([
				[0, 0],
				[w - o, 0],
				[w, h / 2],
				[w - o, h],
				[0, h],
				[o, h / 2],
			]);
		}
		case 'hexagon': {
			const o = ss * adj('adj', 25000);
			return poly([
				[o, 0],
				[w - o, 0],
				[w, h / 2],
				[w - o, h],
				[o, h],
				[0, h / 2],
			]);
		}
		case 'octagon': {
			const o = ss * adj('adj', 29289);
			return poly([
				[o, 0],
				[w - o, 0],
				[w, o],
				[w, h - o],
				[w - o, h],
				[o, h],
				[0, h - o],
				[0, o],
			]);
		}
		case 'plus': {
			const o = ss * adj('adj', 25000);
			return poly(
				[
					[o, 0],
					[w - o, 0],
					[w - o, o],
					[w, o],
					[w, h - o],
					[w - o, h - o],
					[w - o, h],
					[o, h],
					[o, h - o],
					[0, h - o],
					[0, o],
					[o, o],
				].map(([x, y]) => [x, y] as [number, number]),
			);
		}
		case 'rightArrow':
		case 'leftArrow':
		case 'upArrow':
		case 'downArrow': {
			const shaft = adj('adj1', 50000);
			const head = ss * adj('adj2', 50000);
			if (preset === 'rightArrow' || preset === 'leftArrow') {
				const t = (h * (1 - shaft)) / 2;
				const pts: Array<[number, number]> = [
					[0, t],
					[w - head, t],
					[w - head, 0],
					[w, h / 2],
					[w - head, h],
					[w - head, h - t],
					[0, h - t],
				];
				return poly(preset === 'rightArrow' ? pts : pts.map(([x, y]) => [w - x, y]));
			}
			const t = (w * (1 - shaft)) / 2;
			const pts: Array<[number, number]> = [
				[t, h],
				[t, head],
				[0, head],
				[w / 2, 0],
				[w, head],
				[w - t, head],
				[w - t, h],
			];
			return poly(preset === 'upArrow' ? pts : pts.map(([x, y]) => [x, h - y]));
		}
		case 'star5': {
			const pts: Array<[number, number]> = [];
			for (let i = 0; i < 10; i++) {
				const r = i % 2 === 0 ? 0.5 : 0.19;
				const a = -Math.PI / 2 + (i * Math.PI) / 5;
				pts.push([w / 2 + Math.cos(a) * w * r, h / 2 + Math.sin(a) * h * r * 1.05 + h * 0.03]);
			}
			return poly(pts);
		}
		case 'line':
		case 'straightConnector1':
			return `M0 0L${w} ${h}`;
		// Connectors (presetShapeDefinitions.xml): open paths from the top-left to the bottom-right corner.
		case 'bentConnector2':
			return `M0 0L${w} 0L${w} ${h}`;
		case 'bentConnector3': {
			const x1 = w * adj('adj1', 50000);
			return `M0 0L${x1} 0L${x1} ${h}L${w} ${h}`;
		}
		case 'bentConnector4': {
			const x1 = w * adj('adj1', 50000);
			const y2 = h * adj('adj2', 50000);
			return `M0 0L${x1} 0L${x1} ${y2}L${w} ${y2}L${w} ${h}`;
		}
		case 'bentConnector5': {
			const x1 = w * adj('adj1', 50000);
			const y2 = h * adj('adj2', 50000);
			const x3 = w * adj('adj3', 50000);
			return `M0 0L${x1} 0L${x1} ${y2}L${x3} ${y2}L${x3} ${h}L${w} ${h}`;
		}
		case 'curvedConnector2':
			return `M0 0C${w / 2} 0 ${w} ${h / 2} ${w} ${h}`;
		case 'curvedConnector3': {
			const x2 = w * adj('adj1', 50000);
			return `M0 0C${x2 / 2} 0 ${x2} ${h / 4} ${x2} ${h / 2}C${x2} ${(h * 3) / 4} ${(x2 + w) / 2} ${h} ${w} ${h}`;
		}
		case 'curvedConnector4': {
			const x2 = w * adj('adj1', 50000);
			const y4 = h * adj('adj2', 50000);
			return `M0 0C${x2 / 2} 0 ${x2} ${y4 / 2} ${x2} ${y4 / 2}C${x2} ${y4} ${(x2 + w) / 2} ${y4} ${(x2 + w) / 2} ${y4}C${w} ${y4} ${w} ${(y4 + h) / 2} ${w} ${h}`;
		}
		case 'curvedConnector5': {
			const x1 = w * adj('adj1', 50000);
			const y2 = h * adj('adj2', 50000);
			const x3 = w * adj('adj3', 50000);
			return `M0 0C${x1 / 2} 0 ${x1} ${y2 / 2} ${x1} ${y2 / 2}C${x1} ${y2} ${(x1 + x3) / 2} ${y2} ${(x1 + x3) / 2} ${y2}C${x3} ${y2} ${x3} ${(y2 + h) / 2} ${x3} ${(y2 + h) / 2}C${x3} ${h} ${(x3 + w) / 2} ${h} ${w} ${h}`;
		}
		case 'wedgeRectCallout':
		case 'wedgeRoundRectCallout': {
			const tx = w * (0.5 + adj('adj1', -20833));
			const ty = h * (0.5 + adj('adj2', 62500));
			return `M0 0H${w}V${h}H${w * 0.42}L${tx} ${ty}L${w * 0.25} ${h}H0Z`;
		}
		case 'can': {
			const ry = (ss * adj('adj', 25000)) / 2;
			return `M0 ${ry}A${w / 2} ${ry} 0 0 1 ${w} ${ry}V${h - ry}A${w / 2} ${ry} 0 0 1 0 ${h - ry}Z M0 ${ry}A${w / 2} ${ry} 0 0 0 ${w} ${ry}`;
		}
		default:
			return `M0 0H${w}V${h}H0Z`;
	}
}

/** Whether a preset is a line or connector (stroked, never filled). */
export function isLinePreset(preset: string): boolean {
	return /^(line|straightConnector\d?|bentConnector\d|curvedConnector\d)$/.test(preset);
}
