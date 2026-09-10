import { Part } from './Part.mjs';
import { PartName } from '../opc/PartName.mjs';
import { ContentTypes, isStoredUncompressed } from '../opc/ContentTypes.mjs';
import { Namespaces } from './Namespaces.mjs';
import { Docx4JException } from '../opc/exceptions.mjs';

/** A part whose content is bytes (docx4j BinaryPart): images, fonts, embeddings, anything unknown. */
export class BinaryPart extends Part {
  private bytes: Uint8Array | undefined;

  constructor(partName: PartName | string, contentType: string = ContentTypes.OCTET_STREAM, relationshipType: string = '') {
    super(partName, contentType, relationshipType);
  }

  get isLoaded(): boolean {
    return this.bytes !== undefined;
  }

  override get compress(): boolean {
    return !isStoredUncompressed(this.contentType);
  }

  /** The bytes, loaded from the source container on first call. */
  async getBytes(): Promise<Uint8Array> {
    if (this.bytes === undefined) {
      const loaded = await this.loadSourceBytes();
      if (loaded === undefined) throw new Docx4JException(`Part ${this.partName} has no content and no source container`);
      this.bytes = loaded;
    }
    return this.bytes;
  }

  setBytes(bytes: Uint8Array): void {
    this.bytes = bytes;
  }
}

/** An image of any kind (docx4j BinaryPartAbstractImage and its Image*Part subclasses); the kind is the content type. */
export class ImagePart extends BinaryPart {
  constructor(partName: PartName | string, contentType: string) {
    super(partName, contentType, Namespaces.IMAGE);
  }
}

/** An embedded Office package, e.g. an .xlsx behind a chart (docx4j EmbeddedPackagePart). Stored, not deflated. */
export class EmbeddedPackagePart extends BinaryPart {
  constructor(partName: PartName | string, contentType: string) {
    super(partName, contentType, Namespaces.EMBEDDED_PKG);
  }
  override get compress(): boolean {
    return false;
  }
}

/** An OLE object's binary (docx4j OleObjectBinaryPart). */
export class OleObjectBinaryPart extends BinaryPart {
  constructor(partName: PartName | string, contentType: string = ContentTypes.OFFICEDOCUMENT_OLE_OBJECT) {
    super(partName, contentType, Namespaces.OLE_OBJECT);
  }
}

/** An obfuscated embedded font (docx4j ObfuscatedFontPart). */
export class ObfuscatedFontPart extends BinaryPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.OFFICEDOCUMENT_FONT, Namespaces.FONT);
  }
}

/** A TrueType font part (docx4j TrueTypeFontPart). */
export class TrueTypeFontPart extends BinaryPart {
  constructor(partName: PartName | string) {
    super(partName, ContentTypes.TRUETYPE_FONT, Namespaces.FONT);
  }
}

/** An altChunk's content, HTML, RTF, plain text or another docx (docx4j AlternativeFormatInputPart). */
export class AlternativeFormatInputPart extends BinaryPart {
  constructor(partName: PartName | string, contentType: string) {
    super(partName, contentType, Namespaces.AF);
  }
  override get compress(): boolean {
    // docx4j: a WordprocessingML altChunk is a zip already
    return !(this.contentType.includes('wordprocessingml') || this.contentType.includes('ms-word'));
  }
}
