// Document properties parts (docx4j DocPropsCorePart, DocPropsExtendedPart, DocPropsCustomPart).
import type * as core from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_core';
import type * as extended from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_extended';
import type * as custom from '@docx4j/generated-objects-ts/modules/org_docx4j_docProps_custom';
import { XmlPart } from '../XmlPart.mjs';
import { PartName } from '../../opc/PartName.mjs';
import { ContentTypes } from '../../opc/ContentTypes.mjs';
import { Namespaces } from '../Namespaces.mjs';

export class DocPropsCorePart extends XmlPart<core.CoreProperties> {
  constructor(partName: PartName | string = '/docProps/core.xml') {
    super(partName, ContentTypes.PACKAGE_COREPROPERTIES, Namespaces.PROPERTIES_CORE,
      { namespaceURI: 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties', localPart: 'coreProperties' });
  }
}

export class DocPropsExtendedPart extends XmlPart<extended.Properties> {
  constructor(partName: PartName | string = '/docProps/app.xml') {
    super(partName, ContentTypes.OFFICEDOCUMENT_EXTENDEDPROPERTIES, Namespaces.PROPERTIES_EXTENDED,
      { namespaceURI: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties', localPart: 'Properties' });
  }
}

export class DocPropsCustomPart extends XmlPart<custom.Properties> {
  constructor(partName: PartName | string = '/docProps/custom.xml') {
    super(partName, ContentTypes.OFFICEDOCUMENT_CUSTOMPROPERTIES, Namespaces.PROPERTIES_CUSTOM,
      { namespaceURI: 'http://schemas.openxmlformats.org/officeDocument/2006/custom-properties', localPart: 'Properties' });
  }
}
