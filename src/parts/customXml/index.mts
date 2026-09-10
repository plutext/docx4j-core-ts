// Custom XML (docx4j CustomXmlDataStoragePropertiesPart; the data part is in DefaultXmlPart.mts).
import type * as props from '@docx4j/generated-objects-ts/modules/org_docx4j_customXmlProperties';
import { XmlPart } from '../XmlPart.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { Namespaces } from '../Namespaces.mjs';

export const DS_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';

/** `/customXml/itemProps1.xml`: the `ds:datastoreItem` with the itemID and schema references. */
export class CustomXmlDataStoragePropertiesPart extends XmlPart<props.DatastoreItem> {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.OFFICEDOCUMENT_CUSTOMXML_DATASTORAGEPROPERTIES, Namespaces.CUSTOM_XML_DATA_STORAGE_PROPERTIES,
      { namespaceURI: DS_NS, localPart: 'datastoreItem' });
  }
}

export { CustomXmlDataStoragePart } from '../DefaultXmlPart.mjs';
