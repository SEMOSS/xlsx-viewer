import type { enDialogsTools } from '../en/dialogs-tools';
import type { AreaDictionary } from '../types';

/** Spanish UI strings for ../en/dialogs-tools.ts (same keys, same order). */
export const esDialogsTools: AreaDictionary<typeof enDialogsTools> = {
	'dialog.textToColumns.title':
		'Asistente para convertir texto en columnas - paso {step} de {count}',
	'dialog.textToColumns.back': '< Atrás',
	'dialog.textToColumns.next': 'Siguiente >',
	'dialog.textToColumns.finish': 'Finalizar',
	'dialog.textToColumns.messageTitle': 'Microsoft Excel',
	'dialog.textToColumns.multipleSelections':
		'Este comando no se puede usar con selecciones múltiples.',
	'dialog.textToColumns.noData': 'No se seleccionaron datos para analizar.',
	'dialog.textToColumns.replaceData': 'Ya hay datos aquí. ¿Desea reemplazarlos?',

	'dialog.textToColumns.guessedDelimited': 'El asistente estima que sus datos son Delimitados.',
	'dialog.textToColumns.guessedFixed': 'El asistente estima que sus datos son De ancho fijo.',
	'dialog.textToColumns.ifCorrect':
		'Si esto es correcto, elija Siguiente, o bien elija el tipo de datos que mejor los describa.',
	'dialog.textToColumns.originalType': 'Tipo de los datos originales',
	'dialog.textToColumns.chooseType':
		'Elija el tipo de archivo que describa los datos con mayor precisión:',
	'dialog.textToColumns.delimited': 'Delimitados',
	'dialog.textToColumns.delimitedHint':
		'- Caracteres como comas o tabulaciones separan cada campo.',
	'dialog.textToColumns.fixedWidth': 'De ancho fijo',
	'dialog.textToColumns.fixedWidthHint':
		'- Los campos están alineados en columnas con espacios entre uno y otro.',
	'dialog.textToColumns.previewSelected': 'Vista previa de los datos seleccionados:',

	'dialog.textToColumns.delimitersIntro':
		'Esta pantalla le permite establecer los separadores contenidos en los datos. Puede ver cómo cambia el texto en la vista previa.',
	'dialog.textToColumns.delimiters': 'Separadores',
	'dialog.textToColumns.tab': 'Tabulación',
	'dialog.textToColumns.semicolon': 'Punto y coma',
	'dialog.textToColumns.comma': 'Coma',
	'dialog.textToColumns.space': 'Espacio',
	'dialog.textToColumns.other': 'Otro:',
	'dialog.textToColumns.otherDelimiter': 'Otro separador',
	'dialog.textToColumns.consecutive': 'Considerar separadores consecutivos como uno solo',
	'dialog.textToColumns.qualifierLabel': 'Calificador de texto:',
	'dialog.textToColumns.qualifier': 'Calificador de texto',
	// Shown in braces: {none}.
	'dialog.textToColumns.noQualifier': 'ninguno',

	'dialog.textToColumns.fixedIntro':
		'Esta pantalla le permite establecer el ancho de los campos (saltos de columna).',
	'dialog.textToColumns.arrowsHint': 'Las líneas con flechas indican un salto de columna.',
	'dialog.textToColumns.createBreak':
		'Para CREAR una línea de salto, haga clic en la posición deseada.',
	'dialog.textToColumns.deleteBreak':
		'Para ELIMINAR una línea de salto, haga doble clic sobre ella.',
	'dialog.textToColumns.moveBreak': 'Para MOVER una línea de salto, haga clic y arrástrela.',
	'dialog.textToColumns.breakPositions': 'Posiciones de salto:',
	'dialog.textToColumns.breakExample': 'p. ej. 5, 12',
	'dialog.textToColumns.dataPreview': 'Vista previa de los datos',

	'dialog.textToColumns.formatIntro':
		'Esta pantalla le permite seleccionar cada columna y establecer el formato de los datos.',
	'dialog.textToColumns.columnFormat': 'Formato de los datos en columnas',
	'dialog.textToColumns.skip': 'No importar columna (saltar)',
	'dialog.textToColumns.generalHint':
		"'General' convierte los valores numéricos en números, los valores de fecha en fechas y todos los demás valores en texto.",
	'dialog.textToColumns.destination': 'Destino:',
	'dialog.textToColumns.skipColumn': 'Saltar columna',
	'dialog.textToColumns.columnHeader': 'Columna {n}: {format}',

	'dialog.removeDuplicates.intro':
		'Para eliminar valores duplicados, seleccione una o más columnas que contengan duplicados.',
	'dialog.removeDuplicates.selectAll': 'Seleccionar todo',
	'dialog.removeDuplicates.unselectAll': 'Anular selección',
	'dialog.removeDuplicates.hasHeaders': 'Mis datos tienen encabezados',
	'dialog.removeDuplicates.columns': 'Columnas',
	'dialog.removeDuplicates.column': 'Columna {column}',
	'dialog.removeDuplicates.noDuplicates': 'No se encontraron valores duplicados.',
	'dialog.removeDuplicates.removedOne': '1 valor duplicado encontrado y quitado',
	'dialog.removeDuplicates.removedOther': '{count} valores duplicados encontrados y quitados',
	'dialog.removeDuplicates.remainOne': '1 valor único permanece',
	'dialog.removeDuplicates.remainOther': '{count} valores únicos permanecen',
	// {removed} and {remaining} are the two sentences above.
	'dialog.removeDuplicates.summary': '{removed}; {remaining}.',

	'dialog.insertFunction.search': 'Buscar una función:',
	'dialog.insertFunction.searchPlaceholder':
		'Escriba una breve descripción de lo que desea hacer y, a continuación, haga clic en Ir',
	'dialog.insertFunction.go': 'Ir',
	'dialog.insertFunction.category': 'O seleccionar una categoría:',
	'dialog.insertFunction.recommended': 'Recomendadas',
	'dialog.insertFunction.mostRecentlyUsed': 'Usadas recientemente',
	'dialog.insertFunction.all': 'Todo',
	'dialog.insertFunction.selectLabel': 'Seleccionar una función:',
	'dialog.insertFunction.select': 'Seleccionar una función',
	'dialog.insertFunction.noMatch': 'Ninguna función coincide con la búsqueda.',
	'dialog.insertFunction.noFunctions': 'No hay funciones.',
	'dialog.insertFunction.noHelp': 'No hay ayuda disponible.',
};
