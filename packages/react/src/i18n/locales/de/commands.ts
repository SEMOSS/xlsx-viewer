import type { enCommands } from '../en/commands';
import type { AreaDictionary } from '../types';

/** German UI strings for ../en/commands.ts (same keys, same order). */
export const deCommands: AreaDictionary<typeof enCommands> = {
	'message.arrayPart': 'Teile einer Matrix können nicht geändert werden.',
	'message.mergeLosesData':
		'Beim Verbinden von Zellen wird nur der Wert oben links beibehalten, und andere Werte werden verworfen.',
	'message.dialogUnavailable': 'Dieses Dialogfeld ist noch nicht verfügbar.',
	'message.invalidReference': 'Der Bezug ist ungültig.',
	'message.invalidReferenceAt': 'Der Bezug ist ungültig: {ref}',
	'message.numberBetween': 'Geben Sie eine Zahl zwischen {min} und {max} ein.',

	'message.overlappingSelections':
		'Diese Aktion kann nicht für überlappende Auswahlbereiche ausgeführt werden.',
	'message.noPrecedents': 'Die aktive Zelle enthält keine Formel, die auf andere Zellen verweist.',
	'message.noDependents': 'Es gibt keine Formeln, die auf die aktive Zelle verweisen.',
	'message.linkUnsupported': 'Dieser Link kann hier nicht geöffnet werden: {target}',

	'message.sheetNameTaken':
		'Dieser Name ist bereits vergeben. Versuchen Sie es mit einem anderen Namen.',
	'message.sheetNameInvalid':
		'Blattnamen dürfen nicht leer oder länger als 31 Zeichen sein und keines dieser Zeichen enthalten: \\ / ? * [ ] :',

	'undo.typing': 'Eingabe von „{text}“',
	'undo.formula': 'Formel',
	'undo.replace': 'Ersetzen',
	'undo.autoFill': 'AutoAusfüllen',
	'undo.fillDown': 'Nach unten ausfüllen',
	'undo.fillRight': 'Nach rechts ausfüllen',
	'undo.cutPaste': 'Ausschneiden und Einfügen',
	'undo.clearComments': 'Kommentare löschen',
	'undo.formatCells': 'Zellen formatieren',
	'undo.cellStyle': 'Formatvorlage: {name}',
	'undo.insertRows': 'Zeilen einfügen',
	'undo.insertColumns': 'Spalten einfügen',
	'undo.deleteRows': 'Zeilen löschen',
	'undo.deleteColumns': 'Spalten löschen',
	'undo.insertCells': 'Zellen einfügen',
	'undo.deleteCells': 'Zellen löschen',
	'undo.rowHeight': 'Zeilenhöhe',
	'undo.columnWidth': 'Spaltenbreite',
	'undo.clearFilter': 'Filter löschen',
	'undo.showLevel': 'Ebene {level} anzeigen',
	'undo.clearRules': 'Regeln löschen',
	'undo.manageRules': 'Regeln verwalten',
	'undo.insertHyperlink': 'Link einfügen',
	'undo.editHyperlink': 'Link bearbeiten',
	'undo.removeHyperlinks': 'Links entfernen',
	'undo.editName': 'Name bearbeiten',
	'undo.deleteName': 'Name löschen',
	'undo.moveSheet': 'Blatt verschieben',
	'undo.copySheet': 'Blatt kopieren',
	'undo.unhideSheet': 'Blatt einblenden',
	'undo.createTable': 'Tabelle erstellen',
	'undo.tableStyle': 'Tabellenformat',
	'undo.renameTable': 'Tabelle umbenennen',
	'undo.chartStyle': 'Diagrammformatvorlage',
	'undo.insertPicture': 'Bild einfügen',
	'undo.insertChart': 'Diagramm einfügen',
	'undo.deleteObject': 'Objekt löschen',
	'undo.moveObject': 'Objekt verschieben',
	'undo.sizeObject': 'Objektgröße ändern',
	'undo.showNote': 'Notiz anzeigen',
	'undo.hideNote': 'Notiz ausblenden',
	'undo.hideAllNotes': 'Alle Notizen ausblenden',
	'undo.printGridlines': 'Gitternetzlinien drucken',
	'undo.printHeadings': 'Überschriften drucken',
};
