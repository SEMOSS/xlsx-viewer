import type { ComponentType } from 'react';

import type { DialogKind } from '../../shell-context';
import { ConditionalRuleDialog } from './ConditionalRuleDialog';
import { DataValidationDialog } from './DataValidationDialog';
import { DefineNameDialog } from './DefineNameDialog';
import { FormatCellsDialog } from './FormatCellsDialog';
import { HyperlinkDialog } from './HyperlinkDialog';
import { DeleteCellsDialog, InsertCellsDialog } from './InsertDeleteCellsDialog';
import { InsertFunctionDialog } from './InsertFunctionDialog';
import { ManageRulesDialog } from './ManageRulesDialog';
import { NameManagerDialog } from './NameManagerDialog';
import { PageSetupDialog } from './PageSetupDialog';
import { PropertiesDialog } from './PropertiesDialog';
import {
	AlertDialog,
	ProtectSheetDialog,
	ProtectWorkbookDialog,
	UnprotectDialog,
} from './ProtectionDialogs';
import { RemoveDuplicatesDialog } from './RemoveDuplicatesDialog';
import { SortDialog } from './SortDialog';
import { TextToColumnsDialog } from './TextToColumnsDialog';
import type { DialogProps } from './types';
import { UnhideSheetDialog } from './UnhideSheetDialog';
import { ValidationAlertDialog } from './ValidationAlertDialog';

/**
 * Feature dialogs (Format Cells, Insert Function, Name Manager, ...) keyed by
 * kind. DialogHost renders these for kinds it does not implement itself.
 */
export const DIALOG_REGISTRY: Partial<Record<DialogKind, ComponentType<DialogProps>>> = {
	alert: AlertDialog,
	protectSheet: ProtectSheetDialog,
	protectWorkbook: ProtectWorkbookDialog,
	unprotect: UnprotectDialog,
	validationAlert: ValidationAlertDialog,
	conditionalRule: ConditionalRuleDialog,
	dataValidation: DataValidationDialog,
	defineName: DefineNameDialog,
	deleteCells: DeleteCellsDialog,
	formatCells: FormatCellsDialog,
	hyperlink: HyperlinkDialog,
	insertCells: InsertCellsDialog,
	insertFunction: InsertFunctionDialog,
	manageRules: ManageRulesDialog,
	nameManager: NameManagerDialog,
	pageSetup: PageSetupDialog,
	properties: PropertiesDialog,
	removeDuplicates: RemoveDuplicatesDialog,
	sort: SortDialog,
	textToColumns: TextToColumnsDialog,
	unhideSheet: UnhideSheetDialog,
};
