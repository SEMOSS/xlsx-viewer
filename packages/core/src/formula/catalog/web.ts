/** Web functions: URL encoding, XML queries and web service calls. */
import { arg, categoryBuilder } from './define';
import type { FunctionInfo } from './types';

const fn = categoryBuilder('Web');

export const WEB_INFO: readonly FunctionInfo[] = [
	fn('ENCODEURL', 'Percent-encodes a text so it can be used safely inside a URL.', [
		arg('text', 'The text to encode.'),
	]),
	fn('FILTERXML', 'Extracts the values that an XPath expression selects from XML text.', [
		arg('xml', 'A string of valid XML.'),
		arg('xpath', 'An XPath expression that selects the data to return.'),
	]),
	fn('WEBSERVICE', 'Downloads the response of a web service at a URL, as text.', [
		arg('url', 'The address of the web service to call.'),
	]),
];
