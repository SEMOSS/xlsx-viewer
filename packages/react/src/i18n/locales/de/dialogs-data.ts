import type { enDialogsData } from '../en/dialogs-data';
import type { AreaDictionary } from '../types';

/** German UI strings for ../en/dialogs-data.ts (same keys, same order). */
export const deDialogsData: AreaDictionary<typeof enDialogsData> = {
	'dialog.sort.title': 'Sortieren',
	'dialog.sort.addLevel': 'Ebene hinzufügen',
	'dialog.sort.deleteLevel': 'Ebene löschen',
	'dialog.sort.copyLevel': 'Ebene kopieren',
	'dialog.sort.moveUp': 'Nach oben',
	'dialog.sort.moveDown': 'Nach unten',
	'dialog.sort.options': 'Optionen...',
	'dialog.sort.hasHeaders': 'Daten haben Überschriften',

	'dialog.sort.levels': 'Sortierebenen',
	'dialog.sort.level': 'Ebene',
	'dialog.sort.column': 'Spalte',
	'dialog.sort.row': 'Zeile',
	'dialog.sort.sortOn': 'Sortieren anhand von',
	'dialog.sort.order': 'Reihenfolge',
	'dialog.sort.sortBy': 'Sortieren nach',
	'dialog.sort.thenBy': 'Dann nach',
	'dialog.sort.levelColumn': 'Ebene {n}: Spalte',
	'dialog.sort.levelRow': 'Ebene {n}: Zeile',
	'dialog.sort.levelSortOn': 'Ebene {n}: Sortieren anhand von',
	'dialog.sort.levelOrder': 'Ebene {n}: Reihenfolge',
	'dialog.sort.columnLetter': 'Spalte {letter}',
	'dialog.sort.rowNumber': 'Zeile {n}',
	'dialog.sort.cellValues': 'Zellwerte',
	'dialog.sort.cellColor': 'Zellenfarbe',
	'dialog.sort.fontColor': 'Schriftfarbe',
	'dialog.sort.cfIcon': 'Bedingtes Formatierungssymbol',
	'dialog.sort.aToZ': 'A bis Z',
	'dialog.sort.zToA': 'Z bis A',
	'dialog.sort.smallestToLargest': 'Nach Größe (aufsteigend)',
	'dialog.sort.largestToSmallest': 'Nach Größe (absteigend)',
	'dialog.sort.oldestToNewest': 'Vom ältesten zum neuesten',
	'dialog.sort.newestToOldest': 'Vom neuesten zum ältesten',

	'dialog.sort.multipleSelections':
		'Diese Aktion kann nicht für eine Mehrfachauswahl ausgeführt werden.',
	'dialog.sort.nothingToSort':
		'Dies kann nicht auf den ausgewählten Bereich angewendet werden. Wählen Sie eine einzelne Zelle in einem Bereich aus, und versuchen Sie es erneut.',
	'dialog.sort.columnRequired':
		'Für alle Sortierkriterien muss eine Spalte angegeben sein. Überprüfen Sie die ausgewählten Sortierkriterien, und versuchen Sie es erneut.',
	'dialog.sort.duplicateCondition':
		'„{label}“ wird mehrmals nach Werten sortiert. Löschen Sie die doppelte Sortierbedingung, und versuchen Sie es erneut.',

	'dialog.sort.optionsTitle': 'Sortieroptionen',
	'dialog.sort.caseSensitive': 'Groß-/Kleinschreibung beachten',
	'dialog.sort.orientation': 'Ausrichtung',
	'dialog.sort.topToBottom': 'Zeilen sortieren',
	'dialog.sort.leftToRight': 'Spalten sortieren',

	'dialog.dataValidation.title': 'Datenüberprüfung',
	'dialog.dataValidation.settings': 'Einstellungen',
	'dialog.dataValidation.inputMessage': 'Eingabemeldung',
	'dialog.dataValidation.errorAlert': 'Fehlermeldung',
	'dialog.dataValidation.clearAll': 'Alle löschen',

	'dialog.dataValidation.criteria': 'Gültigkeitskriterien',
	'dialog.dataValidation.allow': 'Zulassen:',
	'dialog.dataValidation.ignoreBlank': 'Leere Zellen ignorieren',
	'dialog.dataValidation.data': 'Daten:',
	'dialog.dataValidation.inCellDropdown': 'Zellendropdown',
	'dialog.dataValidation.applyToAll':
		'Diese Änderungen auf alle anderen Zellen mit den gleichen Einstellungen anwenden',

	'dialog.dataValidation.type.none': 'Jeden Wert',
	'dialog.dataValidation.type.whole': 'Ganze Zahl',
	'dialog.dataValidation.type.decimal': 'Dezimal',
	'dialog.dataValidation.type.list': 'Liste',
	'dialog.dataValidation.type.date': 'Datum',
	'dialog.dataValidation.type.time': 'Zeit',
	'dialog.dataValidation.type.textLength': 'Textlänge',
	'dialog.dataValidation.type.custom': 'Benutzerdefiniert',

	'dialog.dataValidation.operator.between': 'zwischen',
	'dialog.dataValidation.operator.notBetween': 'nicht zwischen',
	'dialog.dataValidation.operator.equal': 'gleich',
	'dialog.dataValidation.operator.notEqual': 'ungleich',
	'dialog.dataValidation.operator.greaterThan': 'größer als',
	'dialog.dataValidation.operator.lessThan': 'kleiner als',
	'dialog.dataValidation.operator.greaterThanOrEqual': 'größer oder gleich',
	'dialog.dataValidation.operator.lessThanOrEqual': 'kleiner oder gleich',

	'dialog.dataValidation.box.minimum': 'Minimum:',
	'dialog.dataValidation.box.maximum': 'Maximum:',
	'dialog.dataValidation.box.value': 'Wert:',
	'dialog.dataValidation.box.length': 'Länge:',
	'dialog.dataValidation.box.startDate': 'Anfangsdatum:',
	'dialog.dataValidation.box.endDate': 'Enddatum:',
	'dialog.dataValidation.box.date': 'Datum:',
	'dialog.dataValidation.box.startTime': 'Anfangszeit:',
	'dialog.dataValidation.box.endTime': 'Endzeit:',
	'dialog.dataValidation.box.time': 'Zeit:',
	'dialog.dataValidation.box.source': 'Quelle:',
	'dialog.dataValidation.box.formula': 'Formel:',

	'dialog.dataValidation.showInputMessage': 'Eingabemeldung anzeigen, wenn Zelle ausgewählt wird',
	'dialog.dataValidation.inputMessageGroup':
		'Diese Eingabemeldung anzeigen, wenn Zelle ausgewählt wird:',
	'dialog.dataValidation.titleField': 'Titel:',
	'dialog.dataValidation.inputMessageField': 'Eingabemeldung:',

	'dialog.dataValidation.showErrorAlert':
		'Fehlermeldung anzeigen, wenn ungültige Daten eingegeben wurden',
	'dialog.dataValidation.errorAlertGroup':
		'Diese Fehlermeldung anzeigen, wenn ungültige Daten eingegeben werden:',
	'dialog.dataValidation.alertStyle': 'Typ:',
	'dialog.dataValidation.errorStyle.stop': 'Stopp',
	'dialog.dataValidation.errorStyle.warning': 'Warnung',
	'dialog.dataValidation.errorStyle.information': 'Informationen',
	'dialog.dataValidation.icon.stop': 'Stoppsymbol',
	'dialog.dataValidation.icon.warning': 'Warnsymbol',
	'dialog.dataValidation.icon.information': 'Informationssymbol',
	'dialog.dataValidation.errorMessageField': 'Fehlermeldung:',

	'dialog.dataValidation.mixedQuestion':
		'Die Auswahl enthält mehr als eine Art der Datenüberprüfung. Möchten Sie die aktuellen Einstellungen löschen und fortfahren?',
	'dialog.dataValidation.partialQuestion':
		'Die Auswahl enthält einige Zellen ohne Einstellungen für die Datenüberprüfung. Möchten Sie die Datenüberprüfung auf diese Zellen erweitern?',

	'dialog.hyperlink.insertTitle': 'Link einfügen',
	'dialog.hyperlink.editTitle': 'Link bearbeiten',
	'dialog.hyperlink.removeLink': 'Link entfernen',
	'dialog.hyperlink.linkTo': 'Link zu:',
	'dialog.hyperlink.fileOrWebPage': 'Datei oder Webseite',
	'dialog.hyperlink.placeInDocument': 'Aktuelles Dokument',
	'dialog.hyperlink.emailAddress': 'E-Mail-Adresse',
	'dialog.hyperlink.textToDisplay': 'Anzuzeigender Text:',
	'dialog.hyperlink.selectionInDocument': '<<Auswahl im Dokument>>',
	'dialog.hyperlink.screenTip': 'QuickInfo...',
	'dialog.hyperlink.screenTipText': 'QuickInfo-Text:',

	'dialog.hyperlink.address': 'Adresse:',
	'dialog.hyperlink.cellReference': 'Geben Sie den Zellbezug ein:',
	'dialog.hyperlink.selectPlace': 'Oder wählen Sie eine Stelle in diesem Dokument aus:',
	'dialog.hyperlink.cellReferenceGroup': 'Zellbezug',
	'dialog.hyperlink.definedNamesGroup': 'Definierte Namen',
	'dialog.hyperlink.emailField': 'E-Mail-Adresse:',
	'dialog.hyperlink.subject': 'Betreff:',
};
