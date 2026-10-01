import type { enCommands } from '../en/commands';
import type { AreaDictionary } from '../types';

/** French UI strings for ../en/commands.ts (same keys, same order). */
export const frCommands: AreaDictionary<typeof enCommands> = {
	'message.arrayPart': "Impossible de modifier une partie d'une matrice.",
	'message.mergeLosesData':
		'La fusion des cellules conserve uniquement la valeur supérieure gauche et ignore les autres valeurs.',
	'message.dialogUnavailable': "Cette boîte de dialogue n'est pas encore disponible.",
	'message.invalidReference': 'Référence non valide.',
	'message.invalidReferenceAt': 'Référence non valide\u00a0: {ref}',
	'message.numberBetween': 'Entrez un nombre compris entre {min} et {max}.',

	'message.overlappingSelections':
		'Cette action ne peut pas être effectuée sur des sélections qui se chevauchent.',
	'message.noPrecedents':
		"La cellule active ne contient aucune formule faisant référence à d'autres cellules.",
	'message.noDependents': 'Aucune formule ne fait référence à la cellule active.',
	'message.linkUnsupported': "Impossible d'ouvrir ce lien ici\u00a0: {target}",

	'message.sheetNameTaken': 'Ce nom est déjà utilisé. Essayez-en un autre.',
	'message.sheetNameInvalid':
		'Un nom de feuille ne peut ni être vide, ni dépasser 31 caractères, ni contenir \\\u00a0/\u00a0?\u00a0*\u00a0[\u00a0]\u00a0:',

	'undo.typing': 'Saisie de «\u00a0{text}\u00a0»',
	'undo.formula': 'Formule',
	'undo.replace': 'Remplacer',
	'undo.autoFill': 'Recopie incrémentée',
	'undo.fillDown': 'Recopier vers le bas',
	'undo.fillRight': 'Recopier vers la droite',
	'undo.cutPaste': 'Couper et coller',
	'undo.clearComments': 'Effacer les commentaires',
	'undo.formatCells': 'Format de cellule',
	'undo.cellStyle': 'Style\u00a0: {name}',
	'undo.insertRows': 'Insérer des lignes',
	'undo.insertColumns': 'Insérer des colonnes',
	'undo.deleteRows': 'Supprimer des lignes',
	'undo.deleteColumns': 'Supprimer des colonnes',
	'undo.insertCells': 'Insérer des cellules',
	'undo.deleteCells': 'Supprimer des cellules',
	'undo.rowHeight': 'Hauteur de ligne',
	'undo.columnWidth': 'Largeur de colonne',
	'undo.clearFilter': 'Effacer le filtre',
	'undo.showLevel': 'Afficher le niveau {level}',
	'undo.clearRules': 'Effacer les règles',
	'undo.manageRules': 'Gérer les règles',
	'undo.insertHyperlink': 'Insérer un lien hypertexte',
	'undo.editHyperlink': 'Modifier le lien hypertexte',
	'undo.removeHyperlinks': 'Supprimer les liens hypertexte',
	'undo.editName': 'Modifier le nom',
	'undo.deleteName': 'Supprimer le nom',
	'undo.moveSheet': 'Déplacer la feuille',
	'undo.copySheet': 'Copier la feuille',
	'undo.unhideSheet': 'Afficher la feuille',
	'undo.createTable': 'Créer un tableau',
	'undo.tableStyle': 'Style de tableau',
	'undo.renameTable': 'Renommer le tableau',
	'undo.chartStyle': 'Style du graphique',
	'undo.insertPicture': 'Insérer une image',
	'undo.insertChart': 'Insérer un graphique',
	'undo.deleteObject': "Supprimer l'objet",
	'undo.moveObject': "Déplacer l'objet",
	'undo.sizeObject': "Redimensionner l'objet",
	'undo.showNote': 'Afficher la note',
	'undo.hideNote': 'Masquer la note',
	'undo.hideAllNotes': 'Masquer toutes les notes',
	'undo.printGridlines': 'Imprimer le quadrillage',
	'undo.printHeadings': 'Imprimer les en-têtes',
};
