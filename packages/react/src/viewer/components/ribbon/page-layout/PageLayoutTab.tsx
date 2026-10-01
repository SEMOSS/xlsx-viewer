import { PageSetupGroup } from './PageSetupGroup';
import { ScaleToFitGroup } from './ScaleToFitGroup';
import { SheetOptionsGroup } from './SheetOptionsGroup';

/** Page Layout: page setup, scale to fit and sheet options (between Insert and Formulas). */
export function PageLayoutTab() {
	return (
		<>
			<PageSetupGroup />
			<ScaleToFitGroup />
			<SheetOptionsGroup />
		</>
	);
}
