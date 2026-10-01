import { createContext, useContext } from 'react';

export type InsertChartType =
	| 'column'
	| 'bar'
	| 'line'
	| 'pie'
	| 'doughnut'
	| 'area'
	| 'scatter'
	| 'radar';

/** Viewer features the ribbon can trigger that live outside the command layer. */
export interface RibbonHost {
	insertChart?: (type: InsertChartType) => void;
	insertPicture?: () => void;
	insertTable?: () => void;
	toggleFormulaBar?: () => void;
	formulaBarVisible?: boolean;
	zoomToSelection?: () => void;
	/** Top-left cell currently visible in the body (for freezing at the visible origin). */
	topLeftVisible?: () => { row: number; col: number };
}

export const RibbonHostContext = createContext<RibbonHost>({});

export function useRibbonHost(): RibbonHost {
	return useContext(RibbonHostContext);
}
