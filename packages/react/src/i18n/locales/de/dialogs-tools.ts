import type { enDialogsTools } from '../en/dialogs-tools';
import type { AreaDictionary } from '../types';

/** German UI strings for ../en/dialogs-tools.ts (same keys, same order). */
export const deDialogsTools: AreaDictionary<typeof enDialogsTools> = {
	'dialog.textToColumns.title': 'Textkonvertierungs-Assistent - Schritt {step} von {count}',
	'dialog.textToColumns.back': '< Zurück',
	'dialog.textToColumns.next': 'Weiter >',
	'dialog.textToColumns.finish': 'Fertig stellen',
	'dialog.textToColumns.messageTitle': 'Microsoft Excel',
	'dialog.textToColumns.multipleSelections':
		'Dieser Befehl kann nicht für eine Mehrfachauswahl verwendet werden.',
	'dialog.textToColumns.noData': 'Es wurden keine Daten zum Analysieren ausgewählt.',
	'dialog.textToColumns.replaceData':
		'Hier sind bereits Daten vorhanden. Möchten Sie sie ersetzen?',

	'dialog.textToColumns.guessedDelimited':
		'Der Textkonvertierungs-Assistent hat erkannt, dass Ihre Daten mit Trennzeichen versehen sind.',
	'dialog.textToColumns.guessedFixed':
		'Der Textkonvertierungs-Assistent hat erkannt, dass Ihre Daten eine feste Breite haben.',
	'dialog.textToColumns.ifCorrect':
		'Wenn dies zutrifft, klicken Sie auf „Weiter“, oder wählen Sie den Datentyp aus, der Ihre Daten am besten beschreibt.',
	'dialog.textToColumns.originalType': 'Ursprünglicher Datentyp',
	'dialog.textToColumns.chooseType':
		'Wählen Sie den Dateityp aus, der Ihre Daten am besten beschreibt:',
	'dialog.textToColumns.delimited': 'Getrennt',
	'dialog.textToColumns.delimitedHint':
		'- Zeichen wie Kommas oder Tabstopps trennen die einzelnen Felder.',
	'dialog.textToColumns.fixedWidth': 'Feste Breite',
	'dialog.textToColumns.fixedWidthHint':
		'- Felder sind in Spalten ausgerichtet, mit Leerzeichen zwischen den einzelnen Feldern.',
	'dialog.textToColumns.previewSelected': 'Vorschau der markierten Daten:',

	'dialog.textToColumns.delimitersIntro':
		'In diesem Schritt können Sie die Trennzeichen festlegen, die Ihre Daten enthalten. In der Vorschau unten sehen Sie, wie sich dies auf Ihren Text auswirkt.',
	'dialog.textToColumns.delimiters': 'Trennzeichen',
	'dialog.textToColumns.tab': 'Tabstopp',
	'dialog.textToColumns.semicolon': 'Semikolon',
	'dialog.textToColumns.comma': 'Komma',
	'dialog.textToColumns.space': 'Leerzeichen',
	'dialog.textToColumns.other': 'Andere:',
	'dialog.textToColumns.otherDelimiter': 'Anderes Trennzeichen',
	'dialog.textToColumns.consecutive': 'Aufeinanderfolgende Trennzeichen als ein Zeichen behandeln',
	'dialog.textToColumns.qualifierLabel': 'Textqualifizierer:',
	'dialog.textToColumns.qualifier': 'Textqualifizierer',
	// Shown in braces: {none}.
	'dialog.textToColumns.noQualifier': 'kein',

	'dialog.textToColumns.fixedIntro':
		'In diesem Schritt können Sie Feldbreiten (Spaltenwechsel) festlegen.',
	'dialog.textToColumns.arrowsHint': 'Pfeillinien kennzeichnen einen Spaltenwechsel.',
	'dialog.textToColumns.createBreak':
		'Klicken Sie an die gewünschte Position, um eine Wechsellinie zu ERSTELLEN.',
	'dialog.textToColumns.deleteBreak':
		'Doppelklicken Sie auf die Linie, um eine Wechsellinie zu LÖSCHEN.',
	'dialog.textToColumns.moveBreak':
		'Klicken Sie auf die Linie, und ziehen Sie sie, um eine Wechsellinie zu VERSCHIEBEN.',
	'dialog.textToColumns.breakPositions': 'Positionen der Spaltenwechsel:',
	'dialog.textToColumns.breakExample': 'z. B. 5, 12',
	'dialog.textToColumns.dataPreview': 'Datenvorschau',

	'dialog.textToColumns.formatIntro':
		'In diesem Schritt können Sie jede Spalte markieren und das Datenformat festlegen.',
	'dialog.textToColumns.columnFormat': 'Datenformat der Spalten',
	'dialog.textToColumns.skip': 'Spalte nicht importieren (überspringen)',
	'dialog.textToColumns.generalHint':
		'„Standard“ konvertiert numerische Werte in Zahlen, Datumswerte in Datumsangaben und alle übrigen Werte in Text.',
	'dialog.textToColumns.destination': 'Ziel:',
	'dialog.textToColumns.skipColumn': 'Spalte überspringen',
	'dialog.textToColumns.columnHeader': 'Spalte {n}: {format}',

	'dialog.removeDuplicates.intro':
		'Wählen Sie eine oder mehrere Spalten aus, die Duplikate enthalten, um doppelte Werte zu löschen.',
	'dialog.removeDuplicates.selectAll': 'Alle auswählen',
	'dialog.removeDuplicates.unselectAll': 'Auswahl aufheben',
	'dialog.removeDuplicates.hasHeaders': 'Daten haben Überschriften',
	'dialog.removeDuplicates.columns': 'Spalten',
	'dialog.removeDuplicates.column': 'Spalte {column}',
	'dialog.removeDuplicates.noDuplicates': 'Es wurden keine doppelten Werte gefunden.',
	'dialog.removeDuplicates.removedOne': '1 doppelter Wert gefunden und entfernt',
	'dialog.removeDuplicates.removedOther': '{count} doppelte Werte gefunden und entfernt',
	'dialog.removeDuplicates.remainOne': '1 eindeutiger Wert verbleibt',
	'dialog.removeDuplicates.remainOther': '{count} eindeutige Werte verbleiben',
	// {removed} and {remaining} are the two sentences above.
	'dialog.removeDuplicates.summary': '{removed}; {remaining}.',

	'dialog.insertFunction.search': 'Funktion suchen:',
	'dialog.insertFunction.searchPlaceholder':
		'Beschreiben Sie kurz, was Sie tun möchten, und klicken Sie dann auf „Start“',
	'dialog.insertFunction.go': 'Start',
	'dialog.insertFunction.category': 'Oder Kategorie auswählen:',
	'dialog.insertFunction.recommended': 'Empfohlen',
	'dialog.insertFunction.mostRecentlyUsed': 'Zuletzt verwendet',
	'dialog.insertFunction.all': 'Alle',
	'dialog.insertFunction.selectLabel': 'Funktion auswählen:',
	'dialog.insertFunction.select': 'Funktion auswählen',
	'dialog.insertFunction.noMatch': 'Keine Funktion entspricht Ihrer Suche.',
	'dialog.insertFunction.noFunctions': 'Keine Funktionen.',
	'dialog.insertFunction.noHelp': 'Keine Hilfe verfügbar.',
};
