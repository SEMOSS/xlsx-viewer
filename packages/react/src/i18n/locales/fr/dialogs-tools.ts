import type { enDialogsTools } from '../en/dialogs-tools';
import type { AreaDictionary } from '../types';

/** French UI strings for ../en/dialogs-tools.ts (same keys, same order). */
export const frDialogsTools: AreaDictionary<typeof enDialogsTools> = {
	'dialog.textToColumns.title': 'Assistant Conversion - Étape {step} sur {count}',
	'dialog.textToColumns.back': '< Précédent',
	'dialog.textToColumns.next': 'Suivant >',
	'dialog.textToColumns.finish': 'Terminer',
	'dialog.textToColumns.messageTitle': 'Microsoft Excel',
	'dialog.textToColumns.multipleSelections':
		'Cette commande ne peut pas être utilisée sur des sélections multiples.',
	'dialog.textToColumns.noData': "Aucune donnée à analyser n'a été sélectionnée.",
	'dialog.textToColumns.replaceData':
		'Il y a déjà des données ici. Voulez-vous les remplacer\u00a0?',

	'dialog.textToColumns.guessedDelimited':
		"L'Assistant Texte a déterminé que vos données sont de type Délimité.",
	'dialog.textToColumns.guessedFixed':
		"L'Assistant Texte a déterminé que vos données sont de type Largeur fixe.",
	'dialog.textToColumns.ifCorrect':
		'Si ce choix est correct, cliquez sur Suivant, sinon choisissez le type de données qui décrit le mieux vos données.',
	'dialog.textToColumns.originalType': "Type de données d'origine",
	'dialog.textToColumns.chooseType':
		'Choisissez le type de fichier qui décrit le mieux vos données\u00a0:',
	'dialog.textToColumns.delimited': 'Délimité',
	'dialog.textToColumns.delimitedHint':
		'- Des caractères tels que des virgules ou des tabulations séparent chaque champ.',
	'dialog.textToColumns.fixedWidth': 'Largeur fixe',
	'dialog.textToColumns.fixedWidthHint':
		'- Les champs sont alignés en colonnes, avec des espaces entre chaque champ.',
	'dialog.textToColumns.previewSelected': 'Aperçu des données sélectionnées\u00a0:',

	'dialog.textToColumns.delimitersIntro':
		"Cet écran vous permet de choisir les séparateurs contenus dans vos données. Vous pouvez voir les changements sur votre texte dans l'aperçu ci-dessous.",
	'dialog.textToColumns.delimiters': 'Séparateurs',
	'dialog.textToColumns.tab': 'Tabulation',
	'dialog.textToColumns.semicolon': 'Point-virgule',
	'dialog.textToColumns.comma': 'Virgule',
	'dialog.textToColumns.space': 'Espace',
	'dialog.textToColumns.other': 'Autre\u00a0:',
	'dialog.textToColumns.otherDelimiter': 'Autre séparateur',
	'dialog.textToColumns.consecutive':
		'Interpréter des séparateurs identiques consécutifs comme uniques',
	'dialog.textToColumns.qualifierLabel': 'Identificateur de texte\u00a0:',
	'dialog.textToColumns.qualifier': 'Identificateur de texte',
	// Shown in braces: {none}.
	'dialog.textToColumns.noQualifier': 'aucun',

	'dialog.textToColumns.fixedIntro':
		'Cet écran vous permet de définir la largeur des champs (séparateurs de colonnes).',
	'dialog.textToColumns.arrowsHint': 'Les lignes fléchées indiquent un séparateur de colonnes.',
	'dialog.textToColumns.createBreak':
		"Pour CRÉER une ligne de séparation, cliquez à l'emplacement souhaité.",
	'dialog.textToColumns.deleteBreak':
		'Pour SUPPRIMER une ligne de séparation, double-cliquez dessus.',
	'dialog.textToColumns.moveBreak':
		'Pour DÉPLACER une ligne de séparation, cliquez dessus et faites-la glisser.',
	'dialog.textToColumns.breakPositions': 'Positions des séparateurs\u00a0:',
	'dialog.textToColumns.breakExample': 'par ex. 5, 12',
	'dialog.textToColumns.dataPreview': 'Aperçu de données',

	'dialog.textToColumns.formatIntro':
		'Cet écran vous permet de sélectionner chaque colonne et de définir le format des données.',
	'dialog.textToColumns.columnFormat': 'Format des données en colonne',
	'dialog.textToColumns.skip': 'Colonne non distribuée',
	'dialog.textToColumns.generalHint':
		'«\u00a0Standard\u00a0» convertit les valeurs numériques en nombres, les dates en dates et toutes les autres valeurs en texte.',
	'dialog.textToColumns.destination': 'Destination\u00a0:',
	'dialog.textToColumns.skipColumn': 'Ignorer la colonne',
	'dialog.textToColumns.columnHeader': 'Colonne {n}\u00a0: {format}',

	'dialog.removeDuplicates.intro':
		'Pour supprimer les doublons, sélectionnez une ou plusieurs colonnes contenant des doublons.',
	'dialog.removeDuplicates.selectAll': 'Sélectionner tout',
	'dialog.removeDuplicates.unselectAll': 'Désélectionner tout',
	'dialog.removeDuplicates.hasHeaders': 'Mes données ont des en-têtes',
	'dialog.removeDuplicates.columns': 'Colonnes',
	'dialog.removeDuplicates.column': 'Colonne {column}',
	'dialog.removeDuplicates.noDuplicates': 'Aucune valeur en double trouvée.',
	'dialog.removeDuplicates.removedOne': '1 valeur en double trouvée et supprimée',
	'dialog.removeDuplicates.removedOther': '{count} valeurs en double trouvées et supprimées',
	'dialog.removeDuplicates.remainOne': '1 valeur unique restante',
	'dialog.removeDuplicates.remainOther': '{count} valeurs uniques restantes',
	// {removed} and {remaining} are the two sentences above.
	'dialog.removeDuplicates.summary': '{removed}\u00a0; {remaining}.',

	'dialog.insertFunction.search': 'Recherchez une fonction\u00a0:',
	'dialog.insertFunction.searchPlaceholder':
		'Tapez une brève description de ce que vous voulez faire, puis cliquez sur Rechercher',
	'dialog.insertFunction.go': 'Rechercher',
	'dialog.insertFunction.category': 'Ou sélectionnez une catégorie\u00a0:',
	'dialog.insertFunction.recommended': 'Recommandé',
	'dialog.insertFunction.mostRecentlyUsed': 'Dernières utilisées',
	'dialog.insertFunction.all': 'Tous',
	'dialog.insertFunction.selectLabel': 'Sélectionnez une fonction\u00a0:',
	'dialog.insertFunction.select': 'Sélectionnez une fonction',
	'dialog.insertFunction.noMatch': 'Aucune fonction ne correspond à votre recherche.',
	'dialog.insertFunction.noFunctions': 'Aucune fonction.',
	'dialog.insertFunction.noHelp': 'Aucune aide disponible.',
};
