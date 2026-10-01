import type { enDialogsData } from '../en/dialogs-data';
import type { AreaDictionary } from '../types';

/** Spanish UI strings for ../en/dialogs-data.ts (same keys, same order). */
export const esDialogsData: AreaDictionary<typeof enDialogsData> = {
	'dialog.sort.title': 'Ordenar',
	'dialog.sort.addLevel': 'Agregar nivel',
	'dialog.sort.deleteLevel': 'Eliminar nivel',
	'dialog.sort.copyLevel': 'Copiar nivel',
	'dialog.sort.moveUp': 'Subir',
	'dialog.sort.moveDown': 'Bajar',
	'dialog.sort.options': 'Opciones...',
	'dialog.sort.hasHeaders': 'Mis datos tienen encabezados',

	'dialog.sort.levels': 'Niveles de ordenación',
	'dialog.sort.level': 'Nivel',
	'dialog.sort.column': 'Columna',
	'dialog.sort.row': 'Fila',
	'dialog.sort.sortOn': 'Ordenar según',
	'dialog.sort.order': 'Criterio de ordenación',
	'dialog.sort.sortBy': 'Ordenar por',
	'dialog.sort.thenBy': 'Luego por',
	'dialog.sort.levelColumn': 'Nivel {n}: columna',
	'dialog.sort.levelRow': 'Nivel {n}: fila',
	'dialog.sort.levelSortOn': 'Nivel {n}: ordenar según',
	'dialog.sort.levelOrder': 'Nivel {n}: criterio de ordenación',
	'dialog.sort.columnLetter': 'Columna {letter}',
	'dialog.sort.rowNumber': 'Fila {n}',
	'dialog.sort.cellValues': 'Valores de celda',
	'dialog.sort.cellColor': 'Color de celda',
	'dialog.sort.fontColor': 'Color de fuente',
	'dialog.sort.cfIcon': 'Icono de formato condicional',
	'dialog.sort.aToZ': 'A a Z',
	'dialog.sort.zToA': 'Z a A',
	'dialog.sort.smallestToLargest': 'De menor a mayor',
	'dialog.sort.largestToSmallest': 'De mayor a menor',
	'dialog.sort.oldestToNewest': 'De más antiguos a más recientes',
	'dialog.sort.newestToOldest': 'De más recientes a más antiguos',

	'dialog.sort.multipleSelections': 'Esta acción no funciona con selecciones múltiples.',
	'dialog.sort.nothingToSort':
		'No se puede aplicar al rango seleccionado. Seleccione una sola celda de un rango y vuelva a intentarlo.',
	'dialog.sort.columnRequired':
		'Todos los criterios de ordenación deben tener una columna especificada. Compruebe los criterios de ordenación seleccionados y vuelva a intentarlo.',
	'dialog.sort.duplicateCondition':
		'Se está ordenando {label} por valores más de una vez. Elimine la condición de ordenación duplicada y vuelva a intentarlo.',

	'dialog.sort.optionsTitle': 'Opciones de ordenación',
	'dialog.sort.caseSensitive': 'Distinguir mayúsculas de minúsculas',
	'dialog.sort.orientation': 'Orientación',
	'dialog.sort.topToBottom': 'Ordenar de arriba hacia abajo',
	'dialog.sort.leftToRight': 'Ordenar de izquierda a derecha',

	'dialog.dataValidation.title': 'Validación de datos',
	'dialog.dataValidation.settings': 'Configuración',
	'dialog.dataValidation.inputMessage': 'Mensaje de entrada',
	'dialog.dataValidation.errorAlert': 'Mensaje de error',
	'dialog.dataValidation.clearAll': 'Borrar todos',

	'dialog.dataValidation.criteria': 'Criterio de validación',
	'dialog.dataValidation.allow': 'Permitir:',
	'dialog.dataValidation.ignoreBlank': 'Omitir blancos',
	'dialog.dataValidation.data': 'Datos:',
	'dialog.dataValidation.inCellDropdown': 'Celda con lista desplegable',
	'dialog.dataValidation.applyToAll':
		'Aplicar estos cambios a otras celdas con la misma configuración',

	'dialog.dataValidation.type.none': 'Cualquier valor',
	'dialog.dataValidation.type.whole': 'Número entero',
	'dialog.dataValidation.type.decimal': 'Decimal',
	'dialog.dataValidation.type.list': 'Lista',
	'dialog.dataValidation.type.date': 'Fecha',
	'dialog.dataValidation.type.time': 'Hora',
	'dialog.dataValidation.type.textLength': 'Longitud del texto',
	'dialog.dataValidation.type.custom': 'Personalizada',

	'dialog.dataValidation.operator.between': 'entre',
	'dialog.dataValidation.operator.notBetween': 'no está entre',
	'dialog.dataValidation.operator.equal': 'igual a',
	'dialog.dataValidation.operator.notEqual': 'no igual a',
	'dialog.dataValidation.operator.greaterThan': 'mayor que',
	'dialog.dataValidation.operator.lessThan': 'menor que',
	'dialog.dataValidation.operator.greaterThanOrEqual': 'mayor o igual que',
	'dialog.dataValidation.operator.lessThanOrEqual': 'menor o igual que',

	'dialog.dataValidation.box.minimum': 'Mínimo:',
	'dialog.dataValidation.box.maximum': 'Máximo:',
	'dialog.dataValidation.box.value': 'Valor:',
	'dialog.dataValidation.box.length': 'Longitud:',
	'dialog.dataValidation.box.startDate': 'Fecha inicial:',
	'dialog.dataValidation.box.endDate': 'Fecha final:',
	'dialog.dataValidation.box.date': 'Fecha:',
	'dialog.dataValidation.box.startTime': 'Hora de inicio:',
	'dialog.dataValidation.box.endTime': 'Hora de finalización:',
	'dialog.dataValidation.box.time': 'Hora:',
	'dialog.dataValidation.box.source': 'Origen:',
	'dialog.dataValidation.box.formula': 'Fórmula:',

	'dialog.dataValidation.showInputMessage': 'Mostrar mensaje de entrada al seleccionar la celda',
	'dialog.dataValidation.inputMessageGroup':
		'Cuando se seleccione la celda, mostrar este mensaje de entrada:',
	'dialog.dataValidation.titleField': 'Título:',
	'dialog.dataValidation.inputMessageField': 'Mensaje de entrada:',

	'dialog.dataValidation.showErrorAlert':
		'Mostrar mensaje de error si se introducen datos no válidos',
	'dialog.dataValidation.errorAlertGroup':
		'Cuando se introduzcan datos no válidos, mostrar este mensaje de error:',
	'dialog.dataValidation.alertStyle': 'Estilo:',
	'dialog.dataValidation.errorStyle.stop': 'Detener',
	'dialog.dataValidation.errorStyle.warning': 'Advertencia',
	'dialog.dataValidation.errorStyle.information': 'Información',
	'dialog.dataValidation.icon.stop': 'Icono de detención',
	'dialog.dataValidation.icon.warning': 'Icono de advertencia',
	'dialog.dataValidation.icon.information': 'Icono de información',
	'dialog.dataValidation.errorMessageField': 'Mensaje de error:',

	'dialog.dataValidation.mixedQuestion':
		'La selección contiene más de un tipo de validación. ¿Desea borrar la configuración actual y continuar?',
	'dialog.dataValidation.partialQuestion':
		'La selección contiene algunas celdas sin configuración de validación de datos. ¿Desea extender la validación de datos a estas celdas?',

	'dialog.hyperlink.insertTitle': 'Insertar hipervínculo',
	'dialog.hyperlink.editTitle': 'Editar hipervínculo',
	'dialog.hyperlink.removeLink': 'Quitar vínculo',
	'dialog.hyperlink.linkTo': 'Vincular a:',
	'dialog.hyperlink.fileOrWebPage': 'Archivo o página web existente',
	'dialog.hyperlink.placeInDocument': 'Lugar de este documento',
	'dialog.hyperlink.emailAddress': 'Dirección de correo electrónico',
	'dialog.hyperlink.textToDisplay': 'Texto para mostrar:',
	'dialog.hyperlink.selectionInDocument': '<<Selección en el documento>>',
	'dialog.hyperlink.screenTip': 'Información en pantalla...',
	'dialog.hyperlink.screenTipText': 'Texto de información en pantalla:',

	'dialog.hyperlink.address': 'Dirección:',
	'dialog.hyperlink.cellReference': 'Escriba la referencia de celda:',
	'dialog.hyperlink.selectPlace': 'O seleccione un lugar de este documento:',
	'dialog.hyperlink.cellReferenceGroup': 'Referencia de celda',
	'dialog.hyperlink.definedNamesGroup': 'Nombres definidos',
	'dialog.hyperlink.emailField': 'Dirección de correo electrónico:',
	'dialog.hyperlink.subject': 'Asunto:',
};
