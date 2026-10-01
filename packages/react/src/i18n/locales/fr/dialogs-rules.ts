import type { enDialogsRules } from '../en/dialogs-rules';
import type { AreaDictionary } from '../types';

/** French UI strings for ../en/dialogs-rules.ts (same keys, same order). */
export const frDialogsRules: AreaDictionary<typeof enDialogsRules> = {
	'dialog.names.managerTitle': 'Gestionnaire de noms',
	'dialog.names.new': 'Nouveau...',
	'dialog.names.edit': 'Modifier...',
	'dialog.names.list': 'Noms',
	'dialog.names.name': 'Nom',
	'dialog.names.value': 'Valeur',
	'dialog.names.refersToHeader': 'Fait référence à',
	'dialog.names.scope': 'Étendue',
	'dialog.names.comment': 'Commentaire',
	'dialog.names.workbook': 'Classeur',
	'dialog.names.refersTo': 'Fait référence à\u00a0:',
	'dialog.names.cancelEdit': 'Annuler la modification',
	'dialog.names.commitEdit': 'Valider la modification',
	'dialog.names.commit': 'Valider',
	'dialog.names.saveReference':
		'Voulez-vous enregistrer les modifications apportées à la référence du nom\u00a0?',
	'dialog.names.confirmDelete': 'Voulez-vous vraiment supprimer le nom {name}\u00a0?',

	'dialog.names.newTitle': 'Nouveau nom',
	'dialog.names.editTitle': 'Modifier le nom',
	'dialog.names.nameField': 'Nom\u00a0:',
	'dialog.names.scopeField': 'Étendue\u00a0:',
	'dialog.names.commentField': 'Commentaire\u00a0:',

	'dialog.cfRule.title.greaterThan': 'Supérieur à',
	'dialog.cfRule.title.lessThan': 'Inférieur à',
	'dialog.cfRule.title.between': 'Entre',
	'dialog.cfRule.title.equalTo': 'Égal à',
	'dialog.cfRule.title.textContains': 'Texte qui contient',
	'dialog.cfRule.title.dateOccurring': 'Une date se produisant',
	'dialog.cfRule.title.duplicateValues': 'Valeurs en double',
	'dialog.cfRule.title.top10': '10 valeurs les plus élevées',
	'dialog.cfRule.title.topPercent': '10\u00a0% les plus élevés',
	'dialog.cfRule.title.bottom10': '10 valeurs les moins élevées',
	'dialog.cfRule.title.bottomPercent': '10\u00a0% les moins élevés',
	'dialog.cfRule.title.aboveAverage': 'Valeurs supérieures à la moyenne',
	'dialog.cfRule.title.belowAverage': 'Valeurs inférieures à la moyenne',

	'dialog.cfRule.prompt.greaterThan':
		'Appliquer une mise en forme aux cellules qui sont SUPÉRIEURES À\u00a0:',
	'dialog.cfRule.prompt.lessThan':
		'Appliquer une mise en forme aux cellules qui sont INFÉRIEURES À\u00a0:',
	'dialog.cfRule.prompt.between':
		'Appliquer une mise en forme aux cellules qui sont COMPRISES ENTRE\u00a0:',
	'dialog.cfRule.prompt.equalTo':
		'Appliquer une mise en forme aux cellules qui sont ÉGALES À\u00a0:',
	'dialog.cfRule.prompt.textContains':
		'Appliquer une mise en forme aux cellules qui contiennent le texte\u00a0:',
	'dialog.cfRule.prompt.dateOccurring':
		'Appliquer une mise en forme aux cellules qui contiennent une date\u00a0:',
	'dialog.cfRule.prompt.duplicateValues':
		'Appliquer une mise en forme aux cellules qui contiennent\u00a0:',
	'dialog.cfRule.prompt.top':
		'Appliquer une mise en forme aux cellules qui se classent parmi les PREMIÈRES\u00a0:',
	'dialog.cfRule.prompt.bottom':
		'Appliquer une mise en forme aux cellules qui se classent parmi les DERNIÈRES\u00a0:',
	'dialog.cfRule.prompt.aboveAverage':
		'Appliquer une mise en forme aux cellules qui sont SUPÉRIEURES À LA MOYENNE\u00a0:',
	'dialog.cfRule.prompt.belowAverage':
		'Appliquer une mise en forme aux cellules qui sont INFÉRIEURES À LA MOYENNE\u00a0:',

	'dialog.cfRule.from': 'De',
	'dialog.cfRule.and': 'et',
	'dialog.cfRule.to': 'À',
	'dialog.cfRule.duplicate': 'En double',
	'dialog.cfRule.unique': 'Uniques',
	'dialog.cfRule.values': 'valeurs',
	'dialog.cfRule.forSelectedRange': 'pour la plage sélectionnée',
	'dialog.cfRule.with': 'avec',
	'dialog.cfRule.formatWith': 'Mettre en forme avec',

	'dialog.cfRule.period.yesterday': 'Hier',
	'dialog.cfRule.period.today': "Aujourd'hui",
	'dialog.cfRule.period.tomorrow': 'Demain',
	'dialog.cfRule.period.last7Days': 'Dans les 7 derniers jours',
	'dialog.cfRule.period.lastWeek': 'La semaine dernière',
	'dialog.cfRule.period.thisWeek': 'Cette semaine',
	'dialog.cfRule.period.nextWeek': 'La semaine prochaine',
	'dialog.cfRule.period.lastMonth': 'Le mois dernier',
	'dialog.cfRule.period.thisMonth': 'Ce mois-ci',
	'dialog.cfRule.period.nextMonth': 'Le mois prochain',

	'dialog.cfRule.preset.lightRedFillDarkRedText': 'Remplissage rouge clair avec texte rouge foncé',
	'dialog.cfRule.preset.yellowFillDarkYellowText': 'Remplissage jaune avec texte jaune foncé',
	'dialog.cfRule.preset.greenFillDarkGreenText': 'Remplissage vert avec texte vert foncé',
	'dialog.cfRule.preset.lightRedFill': 'Remplissage rouge clair',
	'dialog.cfRule.preset.redText': 'Texte rouge',
	'dialog.cfRule.preset.redBorder': 'Bordure rouge',

	'dialog.cfRule.enterValue': 'Entrez une valeur ou une formule.',
	'dialog.cfRule.enterText': 'Entrez le texte à rechercher.',
	'dialog.cfRule.rankRange': 'Entrez un nombre entier compris entre {min} et {max}.',

	'dialog.manageRules.title': 'Gestionnaire des règles de mise en forme conditionnelle',
	'dialog.manageRules.showRulesFor': 'Afficher les règles de mise en forme pour\u00a0:',
	'dialog.manageRules.currentSelection': 'Sélection actuelle',
	'dialog.manageRules.thisWorksheet': 'Cette feuille de calcul',
	'dialog.manageRules.rules': 'Règles',
	'dialog.manageRules.deleteRule': 'Supprimer la règle',
	'dialog.manageRules.duplicateRule': 'Dupliquer la règle',
	'dialog.manageRules.moveUp': 'Monter',
	'dialog.manageRules.moveDown': 'Descendre',
	'dialog.manageRules.ruleHeader': "Règle (appliquée dans l'ordre indiqué)",
	'dialog.manageRules.format': 'Format',
	'dialog.manageRules.appliesTo': "S'applique à",
	'dialog.manageRules.stopIfTrue': 'Interrompre si Vrai',
	'dialog.manageRules.noRules': "Aucune règle ne s'applique ici.",
	'dialog.manageRules.appliesToField': "S'applique à\u00a0:",
};
