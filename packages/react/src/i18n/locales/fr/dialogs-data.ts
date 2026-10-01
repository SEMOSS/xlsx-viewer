import type { enDialogsData } from '../en/dialogs-data';
import type { AreaDictionary } from '../types';

/** French UI strings for ../en/dialogs-data.ts (same keys, same order). */
export const frDialogsData: AreaDictionary<typeof enDialogsData> = {
	'dialog.sort.title': 'Trier',
	'dialog.sort.addLevel': 'Ajouter un niveau',
	'dialog.sort.deleteLevel': 'Supprimer un niveau',
	'dialog.sort.copyLevel': 'Copier un niveau',
	'dialog.sort.moveUp': 'Monter',
	'dialog.sort.moveDown': 'Descendre',
	'dialog.sort.options': 'Options...',
	'dialog.sort.hasHeaders': 'Mes données ont des en-têtes',

	'dialog.sort.levels': 'Niveaux de tri',
	'dialog.sort.level': 'Niveau',
	'dialog.sort.column': 'Colonne',
	'dialog.sort.row': 'Ligne',
	'dialog.sort.sortOn': 'Trier sur',
	'dialog.sort.order': 'Ordre',
	'dialog.sort.sortBy': 'Trier par',
	'dialog.sort.thenBy': 'Puis par',
	'dialog.sort.levelColumn': 'Niveau {n}, colonne',
	'dialog.sort.levelRow': 'Niveau {n}, ligne',
	'dialog.sort.levelSortOn': 'Niveau {n}, trier sur',
	'dialog.sort.levelOrder': 'Niveau {n}, ordre',
	'dialog.sort.columnLetter': 'Colonne {letter}',
	'dialog.sort.rowNumber': 'Ligne {n}',
	'dialog.sort.cellValues': 'Valeurs de cellule',
	'dialog.sort.cellColor': 'Couleur de cellule',
	'dialog.sort.fontColor': 'Couleur de police',
	'dialog.sort.cfIcon': 'Icône de mise en forme conditionnelle',
	'dialog.sort.aToZ': 'De A à Z',
	'dialog.sort.zToA': 'De Z à A',
	'dialog.sort.smallestToLargest': 'Du plus petit au plus grand',
	'dialog.sort.largestToSmallest': 'Du plus grand au plus petit',
	'dialog.sort.oldestToNewest': 'Du plus ancien au plus récent',
	'dialog.sort.newestToOldest': 'Du plus récent au plus ancien',

	'dialog.sort.multipleSelections':
		'Cette action ne peut pas être effectuée sur des sélections multiples.',
	'dialog.sort.nothingToSort':
		"Impossible d'appliquer cette commande à la plage sélectionnée. Sélectionnez une seule cellule dans une plage et réessayez.",
	'dialog.sort.columnRequired':
		'Une colonne doit être spécifiée pour tous les critères de tri. Vérifiez les critères de tri sélectionnés, puis réessayez.',
	'dialog.sort.duplicateCondition':
		"{label} est trié par valeurs plus d'une fois. Supprimez la condition de tri en double, puis réessayez.",

	'dialog.sort.optionsTitle': 'Options de tri',
	'dialog.sort.caseSensitive': 'Respecter la casse',
	'dialog.sort.orientation': 'Orientation',
	'dialog.sort.topToBottom': 'Trier du haut vers le bas',
	'dialog.sort.leftToRight': 'Trier de la gauche vers la droite',

	'dialog.dataValidation.title': 'Validation des données',
	'dialog.dataValidation.settings': 'Options',
	'dialog.dataValidation.inputMessage': 'Message de saisie',
	'dialog.dataValidation.errorAlert': "Alerte d'erreur",
	'dialog.dataValidation.clearAll': 'Effacer tout',

	'dialog.dataValidation.criteria': 'Critères de validation',
	'dialog.dataValidation.allow': 'Autoriser\u00a0:',
	'dialog.dataValidation.ignoreBlank': 'Ignorer si vide',
	'dialog.dataValidation.data': 'Données\u00a0:',
	'dialog.dataValidation.inCellDropdown': 'Liste déroulante dans la cellule',
	'dialog.dataValidation.applyToAll':
		'Appliquer ces modifications aux cellules de paramètres identiques',

	'dialog.dataValidation.type.none': 'Tout',
	'dialog.dataValidation.type.whole': 'Nombre entier',
	'dialog.dataValidation.type.decimal': 'Décimal',
	'dialog.dataValidation.type.list': 'Liste',
	'dialog.dataValidation.type.date': 'Date',
	'dialog.dataValidation.type.time': 'Heure',
	'dialog.dataValidation.type.textLength': 'Longueur du texte',
	'dialog.dataValidation.type.custom': 'Personnalisé',

	'dialog.dataValidation.operator.between': 'comprise entre',
	'dialog.dataValidation.operator.notBetween': 'non comprise entre',
	'dialog.dataValidation.operator.equal': 'égale à',
	'dialog.dataValidation.operator.notEqual': 'différente de',
	'dialog.dataValidation.operator.greaterThan': 'supérieure à',
	'dialog.dataValidation.operator.lessThan': 'inférieure à',
	'dialog.dataValidation.operator.greaterThanOrEqual': 'supérieure ou égale à',
	'dialog.dataValidation.operator.lessThanOrEqual': 'inférieure ou égale à',

	'dialog.dataValidation.box.minimum': 'Minimum\u00a0:',
	'dialog.dataValidation.box.maximum': 'Maximum\u00a0:',
	'dialog.dataValidation.box.value': 'Valeur\u00a0:',
	'dialog.dataValidation.box.length': 'Longueur\u00a0:',
	'dialog.dataValidation.box.startDate': 'Date de début\u00a0:',
	'dialog.dataValidation.box.endDate': 'Date de fin\u00a0:',
	'dialog.dataValidation.box.date': 'Date\u00a0:',
	'dialog.dataValidation.box.startTime': 'Heure de début\u00a0:',
	'dialog.dataValidation.box.endTime': 'Heure de fin\u00a0:',
	'dialog.dataValidation.box.time': 'Heure\u00a0:',
	'dialog.dataValidation.box.source': 'Source\u00a0:',
	'dialog.dataValidation.box.formula': 'Formule\u00a0:',

	'dialog.dataValidation.showInputMessage':
		'Afficher le message de saisie lorsque la cellule est sélectionnée',
	'dialog.dataValidation.inputMessageGroup':
		'Quand la cellule est sélectionnée, afficher ce message de saisie\u00a0:',
	'dialog.dataValidation.titleField': 'Titre\u00a0:',
	'dialog.dataValidation.inputMessageField': 'Message de saisie\u00a0:',

	'dialog.dataValidation.showErrorAlert':
		"Afficher une alerte d'erreur après la saisie de données non valides",
	'dialog.dataValidation.errorAlertGroup':
		"Quand l'utilisateur entre des données non valides, afficher cette alerte d'erreur\u00a0:",
	'dialog.dataValidation.alertStyle': 'Style\u00a0:',
	'dialog.dataValidation.errorStyle.stop': 'Arrêt',
	'dialog.dataValidation.errorStyle.warning': 'Avertissement',
	'dialog.dataValidation.errorStyle.information': 'Informations',
	'dialog.dataValidation.icon.stop': 'Icône Arrêt',
	'dialog.dataValidation.icon.warning': 'Icône Avertissement',
	'dialog.dataValidation.icon.information': 'Icône Informations',
	'dialog.dataValidation.errorMessageField': "Message d'erreur\u00a0:",

	'dialog.dataValidation.mixedQuestion':
		'La sélection contient plusieurs types de validation. Voulez-vous effacer les paramètres actuels et continuer\u00a0?',
	'dialog.dataValidation.partialQuestion':
		'La sélection contient des cellules sans paramètres de validation des données. Voulez-vous étendre la validation des données à ces cellules\u00a0?',

	'dialog.hyperlink.insertTitle': 'Insérer un lien hypertexte',
	'dialog.hyperlink.editTitle': 'Modifier le lien hypertexte',
	'dialog.hyperlink.removeLink': 'Supprimer le lien',
	'dialog.hyperlink.linkTo': 'Lier à\u00a0:',
	'dialog.hyperlink.fileOrWebPage': 'Fichier ou page Web existant(e)',
	'dialog.hyperlink.placeInDocument': 'Emplacement dans ce document',
	'dialog.hyperlink.emailAddress': 'Adresse de messagerie',
	'dialog.hyperlink.textToDisplay': 'Texte à afficher\u00a0:',
	'dialog.hyperlink.selectionInDocument': '<<Sélection dans le document>>',
	'dialog.hyperlink.screenTip': 'Info-bulle...',
	'dialog.hyperlink.screenTipText': "Texte de l'info-bulle\u00a0:",

	'dialog.hyperlink.address': 'Adresse\u00a0:',
	'dialog.hyperlink.cellReference': 'Tapez la référence de la cellule\u00a0:',
	'dialog.hyperlink.selectPlace': 'Ou sélectionnez un emplacement dans ce document\u00a0:',
	'dialog.hyperlink.cellReferenceGroup': 'Référence de cellule',
	'dialog.hyperlink.definedNamesGroup': 'Noms définis',
	'dialog.hyperlink.emailField': 'Adresse de messagerie\u00a0:',
	'dialog.hyperlink.subject': 'Objet\u00a0:',
};
