/** Shared in-test docx builders (same structures as the Python stage-1 tests). */
import { strToU8, zipSync, type Zippable } from "fflate";

export const CONTENT_TYPES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
  '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
  '<Default Extension="xml" ContentType="application/xml"/>' +
  '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
  "</Types>";

export const ROOT_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
  "</Relationships>";

export const DOCUMENT_RELS_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
  '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
  "</Relationships>";

export const STYLES_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';

/**
 * word/document.xml with:
 * - P1: the §1 worked example (split runs, proofErr noise) → "Master Services Agreement"
 * - P2: rsid-fragmented runs → "The term is five (5) years from the Effective Date."
 * - T1: a 2×2 table
 * - P3: empty paragraph
 * and a trailing w:sectPr (not a paragraph).
 */
export const DOCUMENT_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
  '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>' +
  '<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>' +
  "<w:r><w:rPr><w:b/></w:rPr><w:t>Master</w:t></w:r>" +
  '<w:r><w:t xml:space="preserve"> Services</w:t></w:r>' +
  '<w:proofErr w:type="spellStart"/>' +
  '<w:r><w:t xml:space="preserve">  Agreement </w:t></w:r>' +
  "</w:p>" +
  '<w:p><w:r w:rsidR="00AB12CD" w:rsidRPr="00AB12CD"><w:t xml:space="preserve">The term is five (5) </w:t></w:r>' +
  '<w:r w:rsidR="00FF00AA"><w:rPr><w:b/></w:rPr><w:t>years from the Effective Date.</w:t></w:r></w:p>' +
  "<w:tbl><w:tblGrid><w:gridCol/><w:gridCol/></w:tblGrid>" +
  "<w:tr><w:tc><w:p><w:r><w:t>Term</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>Value</w:t></w:r></w:p></w:tc></w:tr>" +
  "<w:tr><w:tc><w:p><w:r><w:t>Fee</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>$100</w:t></w:r></w:p></w:tc></w:tr>" +
  "</w:tbl>" +
  "<w:p/>" +
  "<w:sectPr/>" +
  "</w:body></w:document>";

export interface DocxParts {
  [name: string]: string;
}

export const DEFAULT_PARTS: DocxParts = {
  "[Content_Types].xml": CONTENT_TYPES_XML,
  "_rels/.rels": ROOT_RELS_XML,
  "word/document.xml": DOCUMENT_XML,
  "word/_rels/document.xml.rels": DOCUMENT_RELS_XML,
  "word/styles.xml": STYLES_XML,
};

/** Build a tiny .docx in memory, with a deliberately non-zero mtime. */
export function buildDocx(parts: DocxParts = DEFAULT_PARTS): Uint8Array {
  const zippable: Zippable = {};
  for (const [name, xml] of Object.entries(parts)) zippable[name] = strToU8(xml);
  return zipSync(zippable, { level: 6, mtime: new Date(2023, 6, 15, 12, 34, 56) });
}

/** Wrap a body XML fragment into a minimal document part. */
export function docWithBody(bodyXml: string): string {
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">' +
    `<w:body>${bodyXml}</w:body></w:document>`
  );
}
