/*
 * CR-006 section 5: the docx4j oracle for XmlPart.selectObjects.
 *
 * For every .docx fixture, runs each of EXPRESSIONS over the main document part through
 * docx4j's XmlUtils.getJAXBAssociationsForXPath - what getJAXBNodesViaXPath is built on,
 * without its throw where a hit has no object - and records for every hit its path, the kind
 * of node, and the class of the object docx4j associates with it.  test/xpath-parity.test.mjs
 * runs the same expressions through selectObjects, which must return the same hits in the same
 * order, each with an object of the same class.
 *
 *   mvn -q compile exec:java -Dexec.mainClass=org.docx4j.parity.XPathHarness \
 *       -Dfixtures=../fixtures -Dout=../golden/xpath -Ddocx4j.commit=<hash>
 *
 * The binder is made from a fresh marshal of the part's contents.  The part is unmarshalled
 * first, so that MainDocumentPart.getBinder() takes its "contents set, no binder yet" branch,
 * which marshals them and binds that DOM: what selectObjects does on every call, and what
 * docx4j's refreshXmlFirst is for.  Without it getBinder() binds the source XML, which keeps
 * what neither object model has and is not what either side's selection runs over.
 *
 * A fixture whose main part marshals an mc:AlternateContent gets a golden that says so and
 * records no hits (CR-006 section 8 item 3).  docx4j keeps the element and both branches in
 * its tree, so its marshal has them - a text box's paragraphs twice, under the Choice and
 * under the Fallback - while this package resolves the choice before it unmarshals.  The two
 * would be answering about different documents.
 */
package org.docx4j.parity;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.TreeMap;

import jakarta.xml.bind.Binder;
import jakarta.xml.bind.JAXBElement;

import org.docx4j.XmlUtils;
import org.docx4j.jaxb.JAXBAssociation;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.w3c.dom.Attr;
import org.w3c.dom.Element;
import org.w3c.dom.Node;

public class XPathHarness {

	/** Raise when the shape of a golden changes; the test reads it. */
	static final String HARNESS_VERSION = "1";

	static final String MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

	/**
	 * The expressions, in the order a golden records them.  Element selections first, since
	 * those are what the oracle is for: the structural ones (every paragraph, run, text, table
	 * and control), the ones inside a run (fields, tabs, breaks, a picture), the predicates of
	 * CR-006 section 1, and positional ones, which are where an order would show.  Then one
	 * attribute and one text selection, which docx4j's binder associates with nothing and this
	 * package maps to the element that owns them (CR-006 section 2): recorded so that the
	 * departure is measured rather than assumed.  Only prefixes both sides bind by default
	 * (docx4j's NamespacePrefixMappings, the objects package's table).
	 */
	static final String[] EXPRESSIONS = {
		"//w:p",
		"//w:r",
		"//w:t",
		"/w:document/w:body/*",
		"//w:tbl",
		"//w:tr",
		"//w:tc",
		"//w:tbl//w:p",
		"//w:sdt",
		"//w:sdtContent",
		"//w:hyperlink",
		"//w:bookmarkStart",
		"//w:bookmarkEnd",
		"//w:sectPr",
		"//w:pPr",
		"//w:rPr",
		"//w:pStyle",
		"//w:rStyle",
		"//w:numPr",
		"//w:ins",
		"//w:del",
		"//w:delText",
		"//w:fldSimple",
		"//w:fldChar",
		"//w:instrText",
		"//w:tab",
		"//w:br",
		"//w:drawing",
		"//wp:inline",
		"//a:graphic",
		"//pic:pic",
		"//w:footnoteReference",
		"//w:commentRangeStart",
		"//w:proofErr",
		"//w:lastRenderedPageBreak",
		"//w:p[w:pPr/w:numPr]",
		"//w:p[w:pPr/w:pStyle/@w:val='Heading1']",
		"//w:r[w:rPr/w:rFonts]",
		"//w:r[w:rPr/w:b]",
		"//w:sdt[w:sdtPr/w:tag]",
		"//w:p[.//w:t]",
		"//w:p[not(w:r)]",
		"//w:p[last()]",
		"(//w:p)[1]",
		"//w:body/w:p[1]",
		"//w:pStyle/@w:val",
		"//w:t/text()",
	};

	public static void main(String[] args) throws Exception {
		File fixtures = new File(Harness.required("fixtures"));
		File out = new File(Harness.required("out"));
		String commit = System.getProperty("docx4j.commit", "unknown");
		if (!fixtures.isDirectory()) throw new IllegalArgumentException("not a directory: " + fixtures);
		Files.createDirectories(out.toPath());

		List<File> documents = Harness.documents(fixtures);
		System.out.println("docx4j commit " + commit + "; " + documents.size() + " fixtures, "
				+ EXPRESSIONS.length + " expressions -> " + out);

		long started = System.currentTimeMillis();
		int failed = 0;
		for (File document : documents) {
			long t = System.currentTimeMillis();
			try {
				Map<String, Object> golden = golden(document, commit);
				Path file = out.toPath().resolve(Harness.basename(document) + ".json");
				Files.write(file, Json.write(golden).getBytes(StandardCharsets.UTF_8));
				System.out.println(String.format("  %-40s %5d ms  %7d bytes%s",
						document.getName(), System.currentTimeMillis() - t, Files.size(file),
						golden.containsKey("hits") ? "" : "  (excluded: mc:AlternateContent)"));
			} catch (Exception e) {
				failed++;
				System.out.println("  " + document.getName() + " FAILED: " + e);
				e.printStackTrace();
			}
		}
		System.out.println((documents.size() - failed) + " goldens in "
				+ (System.currentTimeMillis() - started) + " ms" + (failed == 0 ? "" : ", " + failed + " FAILED"));
		if (failed > 0) System.exit(1);
	}

	static Map<String, Object> golden(File file, String commit) throws Exception {
		WordprocessingMLPackage pkg = WordprocessingMLPackage.load(file);
		MainDocumentPart mdp = pkg.getMainDocumentPart();
		mdp.getContents();                          // so that getBinder() marshals, rather than binding the source
		Binder<Node> binder = mdp.getBinder();
		Object root = mdp.getJaxbElement();         // after getBinder(), which replaces it with the bound tree
		Element marshalled = (Element) binder.getXMLNode(root);
		int alternateContent = marshalled.getElementsByTagNameNS(MC, "AlternateContent").getLength();

		Map<String, String> prefixes = new TreeMap<String, String>();
		Map<String, Object> hits = new LinkedHashMap<String, Object>();
		if (alternateContent == 0) {
			for (String expression : EXPRESSIONS) {
				List<Object> list = new ArrayList<Object>();
				for (JAXBAssociation association : XmlUtils.getJAXBAssociationsForXPath(binder, root, expression, false)) {
					list.add(hit(association, prefixes));
				}
				hits.put(expression, list);
			}
		}

		Map<String, Object> header = new LinkedHashMap<String, Object>();
		header.put("harnessVersion", HARNESS_VERSION);
		header.put("docx4jCommit", commit);
		header.put("docx4jVersion", Harness.docx4jVersion());
		header.put("docx4jCoreJarSha256", Harness.docx4jCoreJarSha256());
		header.put("date", Instant.now().truncatedTo(ChronoUnit.SECONDS).toString());
		header.put("fixture", file.getName());
		header.put("fixtureBytes", Files.size(file.toPath()));
		header.put("part", mdp.getPartName().getName());
		header.put("alternateContent", alternateContent);
		header.put("prefixes", new LinkedHashMap<String, Object>(prefixes));

		Map<String, Object> golden = new LinkedHashMap<String, Object>();
		golden.put("header", header);
		if (alternateContent == 0) golden.put("hits", hits);
		return golden;
	}

	/** One hit: its path, what kind of node it is, and the class of the object docx4j gives for it. */
	private static Map<String, Object> hit(JAXBAssociation association, Map<String, String> prefixes) {
		Node node = association.getDomNode();
		Map<String, Object> hit = new LinkedHashMap<String, Object>();
		hit.put("path", path(node, prefixes));
		hit.put("node", kind(node));
		hit.put("object", typeName(association.getJaxbObject()));
		return hit;
	}

	private static String kind(Node node) {
		switch (node.getNodeType()) {
			case Node.ELEMENT_NODE: return "element";
			case Node.ATTRIBUTE_NODE: return "attribute";
			case Node.TEXT_NODE: case Node.CDATA_SECTION_NODE: return "text";
			default: return "type " + node.getNodeType();
		}
	}

	/**
	 * The class of an associated object as the objects package spells a TYPE_NAME: the Java
	 * package with its dots as underscores, then the class, nested classes dotted -
	 * {@code org.docx4j.wml.PPrBase$PStyle} is {@code org_docx4j_wml.PPrBase.PStyle}.  A
	 * JAXBElement is unwrapped, as this package never hands out the wrapper (CR-006 section 8
	 * item 2); DOM content docx4j keeps as DOM is {@code dom}; no association is null.
	 */
	static String typeName(Object object) {
		if (object == null) return null;
		Object value = object instanceof JAXBElement ? ((JAXBElement<?>) object).getValue() : object;
		if (value == null) return null;
		if (value instanceof Node) return "dom";
		Class<?> c = value.getClass();
		String pkg = c.getPackageName();
		String name = pkg.isEmpty() ? c.getName() : c.getName().substring(pkg.length() + 1);
		return pkg.replace('.', '_') + "." + name.replace('$', '.');
	}

	/**
	 * The node's path in the form this package's {@code pathOf} writes it: qualified names as
	 * docx4j marshalled them, a position on every step below the root counting siblings of the
	 * same expanded name, {@code /@name} for an attribute and {@code /text()} for text.  Every
	 * prefix a path uses is recorded against its namespace, so the other side can spell its own
	 * paths with docx4j's prefixes whatever its own are.
	 */
	static String path(Node node, Map<String, String> prefixes) {
		if (node.getNodeType() == Node.ATTRIBUTE_NODE) {
			Attr attr = (Attr) node;
			record(attr, prefixes);
			Element owner = attr.getOwnerElement();
			return (owner == null ? "" : path(owner, prefixes)) + "/@" + attr.getNodeName();
		}
		if (node.getNodeType() == Node.TEXT_NODE || node.getNodeType() == Node.CDATA_SECTION_NODE) {
			Node parent = node.getParentNode();
			return (parent == null ? "" : path(parent, prefixes)) + "/text()";
		}
		if (node.getNodeType() != Node.ELEMENT_NODE) return "";
		record(node, prefixes);
		Node parent = node.getParentNode();
		if (parent == null || parent.getNodeType() != Node.ELEMENT_NODE) return "/" + node.getNodeName();
		int index = 0;
		for (Node sibling = parent.getFirstChild(); sibling != null; sibling = sibling.getNextSibling()) {
			if (sibling.getNodeType() == Node.ELEMENT_NODE && same(sibling, node)) {
				index++;
				if (sibling == node) break;
			}
		}
		return path(parent, prefixes) + "/" + node.getNodeName() + "[" + index + "]";
	}

	private static boolean same(Node a, Node b) {
		return a.getLocalName().equals(b.getLocalName())
				&& java.util.Objects.equals(a.getNamespaceURI(), b.getNamespaceURI());
	}

	private static void record(Node node, Map<String, String> prefixes) {
		String prefix = node.getPrefix();
		if (prefix != null && node.getNamespaceURI() != null) prefixes.put(prefix, node.getNamespaceURI());
	}
}
