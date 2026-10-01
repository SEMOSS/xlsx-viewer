export type { XmlDocument, XmlElement, XmlNode, XmlRawNode } from './types';
export { parseXml, parseXmlRoot, XmlParseError } from './parser';
export { serializeElement, serializeXml, XML_DECLARATION } from './serializer';
export {
	decodeOoxmlEscapes,
	decodeXmlEntities,
	encodeOoxmlEscapes,
	escapeXmlAttr,
	escapeXmlText,
} from './escape';
export * from './dom';
export * from './namespaces';
