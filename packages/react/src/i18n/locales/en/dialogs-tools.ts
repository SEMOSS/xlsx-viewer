/**
 * English UI strings. Data tool dialogs: Text to Columns (`dialog.textToColumns.*`), Remove Duplicates
 * (`dialog.removeDuplicates.*`) and Insert Function (`dialog.insertFunction.*`).
 */
export const enDialogsTools = {
	'dialog.textToColumns.title': 'Convert Text to Columns Wizard - Step {step} of {count}',
	'dialog.textToColumns.back': '< Back',
	'dialog.textToColumns.next': 'Next >',
	'dialog.textToColumns.finish': 'Finish',
	'dialog.textToColumns.messageTitle': 'Microsoft Excel',
	'dialog.textToColumns.multipleSelections': 'This command cannot be used on multiple selections.',
	'dialog.textToColumns.noData': 'No data was selected to parse.',
	'dialog.textToColumns.replaceData': "There's already data here. Do you want to replace it?",

	'dialog.textToColumns.guessedDelimited':
		'The Text Wizard has determined that your data is Delimited.',
	'dialog.textToColumns.guessedFixed':
		'The Text Wizard has determined that your data is Fixed Width.',
	'dialog.textToColumns.ifCorrect':
		'If this is correct, choose Next, or choose the data type that best describes your data.',
	'dialog.textToColumns.originalType': 'Original data type',
	'dialog.textToColumns.chooseType': 'Choose the file type that best describes your data:',
	'dialog.textToColumns.delimited': 'Delimited',
	'dialog.textToColumns.delimitedHint': '- Characters such as commas or tabs separate each field.',
	'dialog.textToColumns.fixedWidth': 'Fixed width',
	'dialog.textToColumns.fixedWidthHint':
		'- Fields are aligned in columns with spaces between each field.',
	'dialog.textToColumns.previewSelected': 'Preview of selected data:',

	'dialog.textToColumns.delimitersIntro':
		'This screen lets you set the delimiters your data contains. You can see how your text is affected in the preview below.',
	'dialog.textToColumns.delimiters': 'Delimiters',
	'dialog.textToColumns.tab': 'Tab',
	'dialog.textToColumns.semicolon': 'Semicolon',
	'dialog.textToColumns.comma': 'Comma',
	'dialog.textToColumns.space': 'Space',
	'dialog.textToColumns.other': 'Other:',
	'dialog.textToColumns.otherDelimiter': 'Other delimiter',
	'dialog.textToColumns.consecutive': 'Treat consecutive delimiters as one',
	'dialog.textToColumns.qualifierLabel': 'Text qualifier:',
	'dialog.textToColumns.qualifier': 'Text qualifier',
	// Shown in braces: {none}.
	'dialog.textToColumns.noQualifier': 'none',

	'dialog.textToColumns.fixedIntro': 'This screen lets you set field widths (column breaks).',
	'dialog.textToColumns.arrowsHint': 'Lines with arrows signify a column break.',
	'dialog.textToColumns.createBreak': 'To CREATE a break line, click at the desired position.',
	'dialog.textToColumns.deleteBreak': 'To DELETE a break line, double click on the line.',
	'dialog.textToColumns.moveBreak': 'To MOVE a break line, click and drag it.',
	'dialog.textToColumns.breakPositions': 'Break positions:',
	'dialog.textToColumns.breakExample': 'e.g. 5, 12',
	'dialog.textToColumns.dataPreview': 'Data preview',

	'dialog.textToColumns.formatIntro':
		'This screen lets you select each column and set the Data Format.',
	'dialog.textToColumns.columnFormat': 'Column data format',
	'dialog.textToColumns.skip': 'Do not import column (skip)',
	'dialog.textToColumns.generalHint':
		"'General' converts numeric values to numbers, date values to dates, and all remaining values to text.",
	'dialog.textToColumns.destination': 'Destination:',
	'dialog.textToColumns.skipColumn': 'Skip column',
	'dialog.textToColumns.columnHeader': 'Column {n}: {format}',

	'dialog.removeDuplicates.intro':
		'To delete duplicate values, select one or more columns that contain duplicates.',
	'dialog.removeDuplicates.selectAll': 'Select All',
	'dialog.removeDuplicates.unselectAll': 'Unselect All',
	'dialog.removeDuplicates.hasHeaders': 'My data has headers',
	'dialog.removeDuplicates.columns': 'Columns',
	'dialog.removeDuplicates.column': 'Column {column}',
	'dialog.removeDuplicates.noDuplicates': 'No duplicate values found.',
	'dialog.removeDuplicates.removedOne': '1 duplicate value found and removed',
	'dialog.removeDuplicates.removedOther': '{count} duplicate values found and removed',
	'dialog.removeDuplicates.remainOne': '1 unique value remains',
	'dialog.removeDuplicates.remainOther': '{count} unique values remain',
	// {removed} and {remaining} are the two sentences above.
	'dialog.removeDuplicates.summary': '{removed}; {remaining}.',

	'dialog.insertFunction.search': 'Search for a function:',
	'dialog.insertFunction.searchPlaceholder':
		'Type a brief description of what you want to do and then click Go',
	'dialog.insertFunction.go': 'Go',
	'dialog.insertFunction.category': 'Or select a category:',
	'dialog.insertFunction.recommended': 'Recommended',
	'dialog.insertFunction.mostRecentlyUsed': 'Most Recently Used',
	'dialog.insertFunction.all': 'All',
	'dialog.insertFunction.selectLabel': 'Select a function:',
	'dialog.insertFunction.select': 'Select a function',
	'dialog.insertFunction.noMatch': 'No functions match your search.',
	'dialog.insertFunction.noFunctions': 'No functions.',
	'dialog.insertFunction.noHelp': 'No help available.',
} as const;
