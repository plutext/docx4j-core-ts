/*
 * CR-001 Phase B step 1: the Java parity harness of @docx4j/core-ts.
 *
 * For every .docx fixture, records what docx4j answers - PropertyResolver's effective
 * paragraph, run and table properties, Emulator's list labels and counters,
 * RunFontSelector's document font per character - as one golden JSON file, so that the
 * TypeScript port can be measured against docx4j itself rather than against an opinion.
 *
 *   mvn -q -o exec:java -Dfixtures=../fixtures -Dout=../golden -Ddocx4j.commit=$(git -C ../../../docx4j rev-parse HEAD)
 *
 * The XML in a golden is docx4j's own marshalling (its namespace prefixes, no
 * declaration); the TypeScript side unmarshals it through the objects facade and compares
 * object trees, so neither formatting nor prefixes are part of the contract, only content.
 *
 * Determinism (the goldens are committed and diffed weekly against docx4j's head):
 *  - physical font discovery is turned off, so the IdentityPlusMapper's decisions do not
 *    depend on what is installed on the machine that ran the harness (see FONT_NOTE);
 *  - an mc:AlternateContent is resolved by docx4j's own McSelection, with
 *    docx4j.jaxb.mc.preferChoice set to the prefixes this package understands, so docx4j
 *    walks the branch the TypeScript side resolves on the DOM (see MC_PREFER_CHOICE);
 *  - the only date is in the header, and the weekly workflow ignores it.
 */
package org.docx4j.parity;

import java.io.File;
import java.io.IOException;
import java.math.BigInteger;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;

import org.docx4j.Docx4jProperties;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.jaxb.McSelection;
import org.docx4j.model.PropertyResolver;
import org.docx4j.model.listnumbering.Emulator;
import org.docx4j.model.listnumbering.ListLevel;
import org.docx4j.model.listnumbering.NumberingState;
import org.docx4j.model.listnumbering.NumberingStates;
import org.docx4j.fonts.Mapper;
import org.docx4j.fonts.FontDecision;
import org.docx4j.fonts.PhysicalFont;
import org.docx4j.fonts.RunFontSelector;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.CommentsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.EndnotesPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.FootnotesPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.openpackaging.parts.WordprocessingML.StyleDefinitionsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.CTFtnEdn;
import org.docx4j.wml.Comments;
import org.docx4j.wml.DocDefaults;
import org.docx4j.wml.PPrBase.Ind;
import org.docx4j.wml.Lvl;
import org.docx4j.wml.P;
import org.docx4j.wml.PPr;
import org.docx4j.wml.R;
import org.docx4j.wml.RPr;
import org.docx4j.wml.RunDel;
import org.docx4j.wml.Style;
import org.docx4j.wml.Styles;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Text;

public class Harness {

	/** Raise when the shape of a golden changes; the parity test reads it.  Version 1
	 *  reached NumRef, the numbering counters and the table style chain by reflection;
	 *  docx4j 17.1.1 (CR-001 batch 49) made all four public and version 2 calls them.
	 *  Version 3 drops the harness's own mc:AlternateContent branch selection for
	 *  docx4j's (CR-021 phase 1: TraversalUtil in McMode.READ over McSelection). */
	static final String HARNESS_VERSION = "3";

	/**
	 * The {@code mc:Choice/@Requires} prefixes docx4j is told it understands
	 * ({@code docx4j.jaxb.mc.preferChoice}, docx4j CR-021 phase 1): the prefixes of this
	 * package's {@code UNDERSTOOD_NAMESPACES} ({@code src/opc/mce/understood.mts}) through
	 * the objects package's {@code NAMESPACE_PREFIXES} table.  The TypeScript side takes the
	 * first {@code mc:Choice} whose {@code Requires} namespaces are all in that set, so
	 * naming the same prefixes here is what makes the two walk the same branch - a text box
	 * once, from its {@code wps} Choice, which is what Word draws.
	 *
	 * <p>The property is read on every selection, so it must be set before any walk.  Nine
	 * of the understood namespaces have no prefix in that table (MathML, InkML, the two
	 * Excel mains, the three encryption ones) and SpreadsheetML's binding there is the
	 * default namespace; none of them is nameable here, so a {@code Requires} naming one
	 * would diverge.  No fixture has one.</p>
	 */
	static final String MC_PREFER_CHOICE =
			"a a13cmd a14 a15 a16 a1611 a16svg a18hc adec am3d an18 anam3d b c c14 c15 c16 c16ac "
			+ "c173 cdr cdr14 comp cp cppr cs cx dc dcterms dgm dgm14 dgm1612 ds dsp iact ink16 lc m "
			+ "mc msink o p p13cmd p14 p15 p1510 p159 p16 p166 p1710 p173 p184 pic pic14 pkg prop "
			+ "properties psez pslz psuz pvml r rel sl thm15 v vt w w10 w14 w15 w16cid w16se we wetp "
			+ "wne wp wp14 wp15 wpc wpg wps xdr xdr14 xvml";

	static final String W = Namespaces.NS_WORD12;

	/**
	 * The font environment the IdentityPlusMapper's decisions were made in.  See the
	 * README: the machine's fonts are out and docx4j's jars are in, so a golden made on a
	 * developer's box equals one made in CI.
	 */
	static final String FONT_NOTE =
			"the docx4j font jars alone (docx4j.fonts.discoverPhysicalFonts.enabled=false, "
			+ "docx4j.fonts.discoverJarFonts.enabled=true, the symbol, croscore, crosextra and "
			+ "theme2023 jars on the classpath): CR-016 phase 0c's -Dfidelity.fonts=jars "
			+ "environment, so the mapping does not depend on what this machine has installed";

	/** How much of a paragraph's text is kept, as a human check on the addresses. */
	static final int TEXT_PREVIEW = 80;

	// --------------------------------------------------------------------------- main

	public static void main(String[] args) throws Exception {

		// Before anything constructs a Mapper: its static initialiser is what discovers
		// fonts.  The machine's own fonts are out (they differ from box to box); docx4j's
		// font jars, which the pom pins, are in.  A fresh font cache directory per run, so
		// that a cache left by something else cannot reach a golden either.
		Docx4jProperties.setProperty("docx4j.fonts.discoverPhysicalFonts.enabled", "false");
		Docx4jProperties.setProperty("docx4j.fonts.discoverJarFonts.enabled", "true");
		System.setProperty("docx4j.fonts.fontcache",
				Files.createTempDirectory("docx4j-parity-fontcache").toString());

		// Before any walk: McSelection reads this property on every mc:AlternateContent, and
		// TraversalUtil's default McMode.READ then gives up the one branch docx4j draws.
		Docx4jProperties.setProperty(McSelection.PROPERTY, MC_PREFER_CHOICE);

		File fixtures = new File(required("fixtures"));
		File out = new File(required("out"));
		String commit = System.getProperty("docx4j.commit", "unknown");
		if (!fixtures.isDirectory()) throw new IllegalArgumentException("not a directory: " + fixtures);
		Files.createDirectories(out.toPath());

		List<File> documents = documents(fixtures);
		System.out.println("docx4j commit " + commit + "; " + documents.size() + " fixtures -> " + out);

		long started = System.currentTimeMillis();
		int failed = 0;
		for (File document : documents) {
			long t = System.currentTimeMillis();
			try {
				Map<String, Object> golden = new Harness().golden(document, commit);
				Path file = out.toPath().resolve(basename(document) + ".json");
				Files.write(file, Json.write(golden).getBytes(StandardCharsets.UTF_8));
				System.out.println(String.format("  %-40s %5d ms  %7d bytes",
						document.getName(), System.currentTimeMillis() - t, Files.size(file)));
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

	private static String required(String property) {
		String value = System.getProperty(property);
		if (value == null || value.isEmpty()) {
			throw new IllegalArgumentException("-D" + property + "=... is required");
		}
		return value;
	}

	/** Every .docx directly under the fixtures directory and under its parity/ directory,
	 *  in one sorted order (.pptx, .xlsx and loose .xml are not WordprocessingML). */
	private static List<File> documents(File fixtures) {
		List<File> all = new ArrayList<File>();
		docx(fixtures, all);
		docx(new File(fixtures, "parity"), all);
		all.sort(Comparator.comparing(File::getName));
		return all;
	}

	private static void docx(File dir, List<File> into) {
		File[] files = dir.listFiles();
		if (files == null) return;
		for (File f : files) {
			if (f.isFile() && f.getName().toLowerCase().endsWith(".docx")) into.add(f);
		}
	}

	private static String basename(File f) {
		String n = f.getName();
		return n.substring(0, n.length() - ".docx".length());
	}

	// ----------------------------------------------------------------- per document

	private WordprocessingMLPackage pkg;
	private MainDocumentPart mdp;
	private PropertyResolver resolver;
	private NumberingDefinitionsPart numbering;
	private NumberingStates states;
	private RunFontSelector fontSelector;
	private Mapper mapper;
	private final List<String> notes = new ArrayList<String>();

	Map<String, Object> golden(File file, String commit) throws Exception {

		pkg = WordprocessingMLPackage.load(file);
		mdp = pkg.getMainDocumentPart();
		resolver = mdp.getPropertyResolver();
		numbering = mdp.getNumberingDefinitionsPart();
		states = new NumberingStates();
		mapper = pkg.getFontMapper();     // an IdentityPlusMapper, populated from fontsInUse()
		fontSelector = new RunFontSelector(pkg, org.docx4j.fonts.FontsAnalysis.NO_OP_VISITOR,
				RunFontSelector.RunFontActionType.XHTML);

		Map<String, Object> golden = new LinkedHashMap<String, Object>();
		Map<String, Object> stories = stories();          // before the header: it collects notes
		Map<String, Object> styles = styles();
		List<Object> tables = tables();
		Map<String, Object> fonts = fonts();

		golden.put("header", header(file, commit));
		golden.put("styles", styles);
		golden.put("stories", stories);
		golden.put("tables", tables);
		golden.put("fonts", fonts);
		return golden;
	}

	private Map<String, Object> header(File file, String commit) throws IOException {
		Map<String, Object> header = new LinkedHashMap<String, Object>();
		header.put("harnessVersion", HARNESS_VERSION);
		header.put("docx4jCommit", commit);
		header.put("docx4jVersion", docx4jVersion());
		header.put("docx4jCoreJarSha256", docx4jCoreJarSha256());
		header.put("date", Instant.now().truncatedTo(ChronoUnit.SECONDS).toString());
		header.put("fixture", file.getName());
		header.put("fixtureBytes", Files.size(file.toPath()));
		header.put("mapping", FONT_NOTE);
		header.put("mcPreferChoice", MC_PREFER_CHOICE);
		header.put("notes", new ArrayList<Object>(notes));
		return header;
	}

	/**
	 * The docx4j-core jar this run actually used, by content: {@code docx4jCommit} says
	 * which commit the harness was <em>told</em> it was running against, and one shared
	 * {@code ~/.m2} can hold a SNAPSHOT installed from somewhere else entirely (a
	 * developer's branch, a worktree).  Sixteen hex digits is enough to tell two builds
	 * apart.  The weekly workflow ignores this line when it diffs, since a rebuild of
	 * unchanged sources gives a new jar; a golden that differs for a real reason carries
	 * the new value with it.
	 */
	private static String docx4jCoreJarSha256() {
		try {
			java.net.URL location = WordprocessingMLPackage.class.getProtectionDomain()
					.getCodeSource().getLocation();
			File jar = new File(location.toURI());
			if (!jar.isFile()) return "not a jar: " + jar.getName();
			java.security.MessageDigest digest = java.security.MessageDigest.getInstance("SHA-256");
			byte[] bytes = Files.readAllBytes(jar.toPath());
			StringBuilder hex = new StringBuilder();
			for (byte b : digest.digest(bytes)) hex.append(String.format("%02x", b));
			return hex.substring(0, 16);
		} catch (Exception e) {
			return "unknown";
		}
	}

	/** The docx4j jar's version: its manifest where it has one, else the jar's file name. */
	private static String docx4jVersion() {
		Package p = WordprocessingMLPackage.class.getPackage();
		String version = p == null ? null : p.getImplementationVersion();
		if (version != null) return version;
		try {
			String location = WordprocessingMLPackage.class.getProtectionDomain()
					.getCodeSource().getLocation().getPath();
			return new File(location).getName();
		} catch (Exception e) {
			return "unknown";
		}
	}

	// ---------------------------------------------------------------------- styles

	private Map<String, Object> styles() {

		Map<String, Object> result = new LinkedHashMap<String, Object>();
		StyleDefinitionsPart sdp = mdp.getStyleDefinitionsPart();
		result.put("defaultParagraphStyleId", resolver.getDefaultParagraphStyleId());

		Map<String, Object> defaults = new LinkedHashMap<String, Object>();
		defaults.put("pPr", xml(resolver.getDocumentDefaultPPr(), "pPr", PPr.class));
		defaults.put("rPr", xml(resolver.getDocumentDefaultRPr(), "rPr", RPr.class));
		result.put("documentDefaults", defaults);

		// the docDefaults element as the part states it, for a port which reads it itself
		DocDefaults docDefaults = null;
		Styles styleList = null;
		if (sdp != null) {
			try {
				styleList = sdp.getJaxbElement();
				docDefaults = styleList == null ? null : styleList.getDocDefaults();
			} catch (Exception e) {
				note("styles part: " + e);
			}
		}
		result.put("docDefaults", xml(docDefaults, "docDefaults", DocDefaults.class));

		Map<String, Object> byId = new TreeMap<String, Object>();
		if (styleList != null) {
			for (Style style : styleList.getStyle()) {
				if (style.getStyleId() == null) continue;
				Map<String, Object> entry = new LinkedHashMap<String, Object>();
				entry.put("type", style.getType());
				entry.put("basedOn", style.getBasedOn() == null ? null : style.getBasedOn().getVal());
				entry.put("isDefault", style.isDefault());
				try {
					entry.put("effectivePPr", xml(resolver.getEffectivePPr(style.getStyleId()), "pPr", PPr.class));
				} catch (Exception e) {
					entry.put("effectivePPr", error(e));
				}
				try {
					entry.put("effectiveRPr", xml(resolver.getEffectiveRPr(style.getStyleId()), "rPr", RPr.class));
				} catch (Exception e) {
					entry.put("effectiveRPr", error(e));
				}
				byId.put(style.getStyleId(), entry);
			}
		}
		result.put("byId", byId);
		return result;
	}

	/** A resolution that threw, in place of the XML: cyclic styles are a legitimate golden. */
	private Map<String, Object> error(Exception e) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("exception", e.getClass().getName());
		m.put("message", e.getMessage());
		note(e.getClass().getSimpleName() + ": " + e.getMessage());
		return m;
	}

	private void note(String text) {
		if (!notes.contains(text)) notes.add(text);
	}

	// --------------------------------------------------------------------- stories

	private Map<String, Object> stories() {

		Map<String, Object> stories = new LinkedHashMap<String, Object>();
		stories.put("main", story(mdp, mdp.getContent()));

		Map<String, HeaderPart> headers = new TreeMap<String, HeaderPart>();
		Map<String, FooterPart> footers = new TreeMap<String, FooterPart>();
		RelationshipsPart rels = mdp.getRelationshipsPart();
		if (rels != null) {
			for (Relationship r : rels.getRelationships().getRelationship()) {
				Part part = rels.getPart(r);
				if (part instanceof HeaderPart) headers.put(r.getId(), (HeaderPart) part);
				else if (part instanceof FooterPart) footers.put(r.getId(), (FooterPart) part);
			}
		}
		for (Map.Entry<String, HeaderPart> e : headers.entrySet()) {
			stories.put("header:" + e.getKey(), story(e.getValue(), e.getValue().getContent()));
		}
		for (Map.Entry<String, FooterPart> e : footers.entrySet()) {
			stories.put("footer:" + e.getKey(), story(e.getValue(), e.getValue().getContent()));
		}

		FootnotesPart footnotes = mdp.getFootnotesPart();
		if (footnotes != null) {
			List<Object> content = new ArrayList<Object>();
			try {
				for (CTFtnEdn note : footnotes.getJaxbElement().getFootnote()) content.addAll(note.getContent());
			} catch (Exception e) {
				note("footnotes: " + e);
			}
			stories.put("footnotes", story(footnotes, content));
		}
		EndnotesPart endnotes = mdp.getEndNotesPart();
		if (endnotes != null) {
			List<Object> content = new ArrayList<Object>();
			try {
				for (CTFtnEdn note : endnotes.getJaxbElement().getEndnote()) content.addAll(note.getContent());
			} catch (Exception e) {
				note("endnotes: " + e);
			}
			stories.put("endnotes", story(endnotes, content));
		}
		CommentsPart comments = mdp.getCommentsPart();
		if (comments != null) {
			List<Object> content = new ArrayList<Object>();
			try {
				for (Comments.Comment c : comments.getJaxbElement().getComment()) content.addAll(c.getContent());
			} catch (Exception e) {
				note("comments: " + e);
			}
			stories.put("comments", story(comments, content));
		}
		return stories;
	}

	private Map<String, Object> story(Part part, List<Object> content) {
		List<Object> paragraphs = new ArrayList<Object>();
		Walk walk = new Walk(paragraphs, states.forPart(part));
		walk.children(content, 0);
		Map<String, Object> story = new LinkedHashMap<String, Object>();
		story.put("part", part.getPartName().getName());
		story.put("paragraphs", paragraphs);
		return story;
	}

	/**
	 * A story's paragraphs in document order, descending into tables, content controls,
	 * hyperlinks, revisions and text boxes, as TraversalUtil walks.  A text box's
	 * paragraphs stay in this story's list (they are in the part, in document order) but
	 * count in a story of their own, which is docx4j's rule: CR-014 probe P7, applied by
	 * AbstractWmlConversionContext.enterTextBox.
	 */
	private final class Walk {

		private final List<Object> into;
		private final NumberingState state;

		Walk(List<Object> into, NumberingState state) {
			this.into = into;
			this.state = state;
		}

		void children(List<?> children, int depth) {
			if (children == null || depth > 60) return;
			for (Object raw : children) {
				Object o = XmlUtils.unwrap(raw);
				if (o == null) continue;
				if (o instanceof P) {
					into.add(paragraph((P) o, into.size(), state));
					children(childrenOf(o), depth + 1);
				} else if (isTextBox(o)) {
					new Walk(into, states.newStory()).children(childrenOf(o), depth + 1);
				} else {
					children(childrenOf(o), depth + 1);
				}
			}
		}
	}

	/**
	 * The node whose children are a text box's content, and so begin a story of their
	 * own: the VML {@code v:textbox}, the {@code w14:wsp} shape, and the
	 * {@code wp:inline} or {@code wp:anchor} that TraversalUtil descends straight through
	 * (graphic, graphicData, {@code wps:wsp}, {@code wps:txbx}) into a
	 * {@code w:txbxContent}.
	 */
	private static boolean isTextBox(Object o) {
		String name = o.getClass().getName();
		if (name.equals("org.docx4j.vml.CTTextbox")) return true;
		if (name.endsWith(".wordprocessingShape.CTWordprocessingShape")) return true;
		if (o instanceof org.docx4j.dml.wordprocessingDrawing.Inline) {
			return hasTextBox(((org.docx4j.dml.wordprocessingDrawing.Inline) o).getGraphic());
		}
		if (o instanceof org.docx4j.dml.wordprocessingDrawing.Anchor) {
			return hasTextBox(((org.docx4j.dml.wordprocessingDrawing.Anchor) o).getGraphic());
		}
		if (o instanceof org.docx4j.dml.GraphicData) {
			return hasTextBox((org.docx4j.dml.GraphicData) o);
		}
		return false;
	}

	private static boolean hasTextBox(org.docx4j.dml.Graphic graphic) {
		return graphic != null && hasTextBox(graphic.getGraphicData());
	}

	private static boolean hasTextBox(org.docx4j.dml.GraphicData graphicData) {
		if (graphicData == null) return false;
		org.docx4j.com.microsoft.schemas.office.word.x2010.wordprocessingShape.CTWordprocessingShape wsp =
				graphicData.getWordprocessingShape();
		return wsp != null && wsp.getTxbx() != null && wsp.getTxbx().getTxbxContent() != null;
	}

	/**
	 * The children to walk: {@link TraversalUtil}'s, in its default {@link
	 * org.docx4j.jaxb.McMode#READ} - so an {@code mc:AlternateContent} gives up the one
	 * branch docx4j draws, {@link McSelection#selectedBranch}, and the walk descends into
	 * that {@code mc:Choice} or {@code mc:Fallback} on the next turn.
	 *
	 * <p>Which branch that is comes from {@link #MC_PREFER_CHOICE}, set in {@code main}:
	 * the first {@code mc:Choice} whose {@code Requires} prefixes this package understands,
	 * else the {@code mc:Fallback}.  So a text box appears once, from its {@code wps}
	 * Choice, which is what Word draws and what the TypeScript side resolves on the DOM
	 * before unmarshalling.  Before docx4j CR-021 phase 1 a walk saw the choices
	 * <em>and</em> the fallback, which put a text box's paragraphs in a golden twice, and
	 * harness versions 1 and 2 took the first Choice themselves.</p>
	 */
	private static List<Object> childrenOf(Object o) {
		return TraversalUtil.getChildrenImpl(o);
	}

	private Map<String, Object> paragraph(P p, int index, NumberingState state) {

		PPr pPr = p.getPPr();
		Map<String, Object> para = new LinkedHashMap<String, Object>();
		para.put("index", index);
		para.put("paraId", p.getParaId());
		para.put("pStyle", pPr == null || pPr.getPStyle() == null ? null : pPr.getPStyle().getVal());
		para.put("text", preview(textOf(p)));
		try {
			para.put("effectivePPr", xml(resolver.getEffectivePPr(pPr), "pPr", PPr.class));
		} catch (Exception e) {
			para.put("effectivePPr", error(e));
		}
		try {
			para.put("paragraphMarkRPr", xml(resolver.getEffectiveParagraphMarkRPr(pPr), "rPr", RPr.class));
		} catch (Exception e) {
			para.put("paragraphMarkRPr", error(e));
		}
		para.put("numbering", numbering(pPr, state));
		para.put("runs", runs(p, pPr));
		return para;
	}

	// ------------------------------------------------------------------- numbering

	private Object numbering(PPr pPr, NumberingState state) {

		if (numbering == null || pPr == null) return null;

		Emulator.NumRef ref = numRef(pPr);
		Emulator.ResultTriple result;
		try {
			result = Emulator.getNumber(pkg, pPr, state);
		} catch (Exception e) {
			Map<String, Object> m = new LinkedHashMap<String, Object>();
			m.put("error", e.toString());
			note("Emulator: " + e);
			return m;
		}
		if (result == null) return null;   // not numbered

		Map<String, Object> num = new LinkedHashMap<String, Object>();
		num.put("numString", result.getNumString());
		num.put("isBullet", result.getBullet() != null);
		num.put("numFont", result.getNumFont());
		num.put("ind", xml(result.getIndent(), "ind", Ind.class));
		num.put("ilvl", ref == null ? null : ref.ilvl);
		num.put("numId", ref == null ? null : ref.numId);
		num.put("labelRPr", xml(result.getLabelRPr(), "rPr", RPr.class));
		num.put("lvl", xml(result.getLvl(), "lvl", Lvl.class));
		num.put("indResolved", indResolved(ref));
		num.put("numRef", numRefJson(ref));
		num.put("stateAfter", stateOf(state));
		return num;
	}

	/** {@code NumberingDefinitionsPart.getInd}: the level's indent with the linked
	 *  paragraph style followed, which is the form the port implements as
	 *  {@code Emulator.getInd}. */
	private Object indResolved(Emulator.NumRef ref) {
		if (ref == null || ref.notNumbered) return null;
		try {
			return xml(numbering.getInd(ref.numId, ref.ilvl), "ind", Ind.class);
		} catch (Exception e) {
			note("getInd: " + e);
			return null;
		}
	}

	private static Map<String, Object> numRefJson(Emulator.NumRef ref) {
		if (ref == null) return null;
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("numId", ref.numId);
		m.put("ilvl", ref.ilvl);
		m.put("direct", ref.direct);
		m.put("notNumbered", ref.notNumbered);
		m.put("reason", ref.reason);
		return m;
	}

	/**
	 * Where docx4j resolves this paragraph's numbering to, before any counting:
	 * {@code Emulator.numRefFor}, which takes no number and touches no state, so it can be
	 * asked before {@code getNumber} increments.  Never null.
	 */
	private Emulator.NumRef numRef(PPr pPr) {
		try {
			return Emulator.numRefFor(pkg, pPr);
		} catch (Exception e) {
			note("Emulator.numRefFor: " + e);
			return null;
		}
	}

	/**
	 * The story's counters after this paragraph: each abstract list's level, the value it
	 * stands at, whether it has been used and whether a shallower level left it pending a
	 * reset, and the w:num start overrides already spent.  Sorted, since the underlying
	 * maps are hashed.
	 */
	private Object stateOf(NumberingState state) {
		try {
			Map<String, Object> result = new LinkedHashMap<String, Object>();
			Map<String, Object> counters = new TreeMap<String, Object>();
			for (Map.Entry<String, ListLevel.Counter> e : state.counters().entrySet()) {
				ListLevel.Counter counter = e.getValue();
				Map<String, Object> c = new LinkedHashMap<String, Object>();
				BigInteger value = counter.getCurrentValue();
				c.put("value", value == null ? null : value.toString());
				c.put("encounteredAlready", counter.isEncounteredAlready());
				c.put("resetPending", counter.isResetPending());
				counters.put(e.getKey(), c);
			}
			result.put("counters", counters);
			result.put("startOverridesApplied",
					new ArrayList<Object>(new TreeSet<String>(state.startOverridesApplied())));
			return result;
		} catch (Exception e) {
			note("NumberingState: " + e);
			return null;
		}
	}

	// ------------------------------------------------------------------------ runs

	private List<Object> runs(P p, PPr pPr) {
		List<Object> runs = new ArrayList<Object>();
		collectRuns(p.getContent(), false, runs, pPr, 0);
		return runs;
	}

	/** Every w:r of the paragraph in order, those inside w:ins, w:hyperlink, w:sdt and
	 *  w:smartTag included; a run inside w:del is flagged rather than dropped.  Nested
	 *  paragraphs (a text box's) and their runs belong to those paragraphs. */
	private void collectRuns(List<?> children, boolean deleted, List<Object> into, PPr pPr, int depth) {
		if (children == null || depth > 60) return;
		for (Object raw : children) {
			Object o = XmlUtils.unwrap(raw);
			if (o == null) continue;
			if (o instanceof R) {
				into.add(run((R) o, pPr, deleted));
			} else if (o instanceof P || isTextBox(o)) {
				continue;
			} else if (o instanceof RunDel) {
				collectRuns(childrenOf(o), true, into, pPr, depth + 1);
			} else {
				collectRuns(childrenOf(o), deleted, into, pPr, depth + 1);
			}
		}
	}

	private Map<String, Object> run(R r, PPr pPr, boolean deleted) {

		Map<String, Object> run = new LinkedHashMap<String, Object>();
		String text = textOf(r);
		run.put("text", text);
		run.put("rStyle", r.getRPr() == null || r.getRPr().getRStyle() == null
				? null : r.getRPr().getRStyle().getVal());
		if (deleted) run.put("deleted", Boolean.TRUE);

		RPr effective = null;
		try {
			effective = resolver.getEffectiveRPr(r.getRPr(), pPr);
			run.put("effectiveRPr", xml(effective, "rPr", RPr.class));
		} catch (Exception e) {
			run.put("effectiveRPr", error(e));
		}
		run.put("fontSpans", fontSpans(pPr, effective, text));
		return run;
	}

	/**
	 * The run's text folded into maximal spans of one document font:
	 * {@code RunFontSelector.documentFontFor} per code point, with the effective rPr
	 * passed in resolved (the {@code rPrIsEffective} overload), which is how
	 * FontsAnalysis walks a document.
	 */
	private List<Object> fontSpans(PPr pPr, RPr effective, String text) {

		List<Object> spans = new ArrayList<Object>();
		if (text == null || text.isEmpty() || effective == null) return spans;

		boolean bold = flag(effective.getB());
		boolean italic = flag(effective.getI());
		boolean cs = flag(effective.getCs());
		boolean rtl = flag(effective.getRtl());

		Map<Integer, String> memo = new HashMap<Integer, String>();
		StringBuilder current = new StringBuilder();
		String currentFont = null;
		for (int i = 0; i < text.length(); ) {
			int cp = text.codePointAt(i);
			int width = Character.charCount(cp);
			String font = memo.get(cp);
			if (font == null) {
				try {
					font = fontSelector.documentFontFor(pPr, effective, cp, true);
				} catch (Exception e) {
					note("documentFontFor: " + e);
					font = "";
				}
				if (font == null) font = "";
				memo.put(cp, font);
			}
			if (currentFont != null && !currentFont.equals(font)) {
				spans.add(span(current.toString(), currentFont, bold, italic, cs, rtl));
				current.setLength(0);
			}
			currentFont = font;
			current.append(text, i, i + width);
			i += width;
		}
		if (current.length() > 0) {
			spans.add(span(current.toString(), currentFont, bold, italic, cs, rtl));
		}
		return spans;
	}

	private static Map<String, Object> span(String text, String font, boolean bold, boolean italic,
			boolean cs, boolean rtl) {
		Map<String, Object> span = new LinkedHashMap<String, Object>();
		span.put("text", text);
		span.put("documentFont", font.isEmpty() ? null : font);
		span.put("bold", bold);
		span.put("italic", italic);
		span.put("cs", cs);
		span.put("rtl", rtl);
		return span;
	}

	private static boolean flag(org.docx4j.wml.BooleanDefaultTrue b) {
		return b != null && b.isVal();
	}

	// ---------------------------------------------------------------------- tables

	private List<Object> tables() {
		List<Object> tables = new ArrayList<Object>();
		collectTables(mdp.getContent(), tables, 0);
		return tables;
	}

	private void collectTables(List<?> children, List<Object> into, int depth) {
		if (children == null || depth > 60) return;
		for (Object raw : children) {
			Object o = XmlUtils.unwrap(raw);
			if (o == null) continue;
			if (o instanceof Tbl) into.add(table((Tbl) o, into.size()));
			collectTables(childrenOf(o), into, depth + 1);
		}
	}

	private Map<String, Object> table(Tbl tbl, int index) {
		Map<String, Object> table = new LinkedHashMap<String, Object>();
		table.put("index", index);
		table.put("tblStyle", tbl.getTblPr() == null || tbl.getTblPr().getTblStyle() == null
				? null : tbl.getTblPr().getTblStyle().getVal());
		try {
			table.put("effectiveTableStyle", xml(resolver.getEffectiveTableStyle(tbl.getTblPr()), "style", Style.class));
		} catch (Exception e) {
			table.put("effectiveTableStyle", error(e));
		}
		table.put("reachesDefaultTableStyle", reachesDefaultTableStyle(tbl));
		return table;
	}

	/**
	 * Whether Word's built-in Normal Table underlies this table: its style chain is empty
	 * or reaches the document's default table style (CR-015 phase 4).  docx4j 17.1.1
	 * exposes the flag {@code getEffectiveTableStyle} decides by; version 1 of this
	 * harness recomputed it from the private style chain.
	 */
	private Object reachesDefaultTableStyle(Tbl tbl) {
		try {
			return resolver.reachesDefaultTableStyle(tbl.getTblPr());
		} catch (Exception e) {
			note("reachesDefaultTableStyle: " + e);
			return null;
		}
	}

	// ----------------------------------------------------------------------- fonts

	private Map<String, Object> fonts() {

		Map<String, Object> fonts = new LinkedHashMap<String, Object>();
		// The theme part and the default font it yields: what every theme reference and
		// every run with no rFonts resolves through, and the one value whose absence would
		// change a whole document's fontSpans, so it is recorded rather than implied.
		fonts.put("themePart", mdp.getThemePart() == null ? null
				: mdp.getThemePart().getPartName().getName());
		fonts.put("defaultFont", fontSelector.getDefaultFont());
		Set<String> inUse = new TreeSet<String>();
		try {
			inUse.addAll(mdp.fontsInUse());
		} catch (Exception e) {
			note("fontsInUse: " + e);
		}
		fonts.put("fontsInUse", new ArrayList<Object>(inUse));

		Set<String> styles = new TreeSet<String>();
		try {
			styles.addAll(mdp.getStylesInUse());
		} catch (Exception e) {
			note("stylesInUse: " + e);
		}
		fonts.put("stylesInUse", new ArrayList<Object>(styles));

		Map<String, Object> mapping = new TreeMap<String, Object>();
		for (String name : inUse) {
			Map<String, Object> decision = new LinkedHashMap<String, Object>();
			FontDecision d = null;
			try {
				d = mapper.getDecision(name);
			} catch (Exception e) {
				note("getDecision: " + e);
			}
			decision.put("source", d == null || d.getSource() == null ? null : d.getSource().name());
			decision.put("via", d == null ? null : d.getVia());
			PhysicalFont physical = d == null ? null : d.getPhysicalFont();
			decision.put("physicalFont", physical == null ? null : physical.getName());
			mapping.put(name, decision);
		}
		fonts.put("mapping", mapping);
		fonts.put("mappingNote", FONT_NOTE);
		return fonts;
	}

	// ------------------------------------------------------------------- utilities

	/**
	 * docx4j's marshalling of one property element: its prefixes, no XML declaration, not
	 * pretty printed (the TypeScript side compares object trees, not text).
	 *
	 * <p>The namespace declarations that no element or attribute of the fragment uses are
	 * dropped first.  docx4j's prefix mapper pre-declares all ninety-odd Office
	 * namespaces on whatever it marshals, which is right for a part and absurd for a
	 * {@code w:ind}: it was 3 kB of {@code xmlns:} on every one of the tens of thousands
	 * of fragments in these goldens, twenty times the content.  Dropping them changes
	 * nothing that is compared, since the comparison unmarshals.  ({@code XmlUtils}'s own
	 * {@code docx4j.jaxb.marshal.canonicalize} would do it through Santuario's
	 * canonicalizer, which is not on docx4j-core's classpath.)</p>
	 */
	private String xml(Object o, String localName, Class<?> declaredType) {
		if (o == null) return null;
		try {
			org.w3c.dom.Document doc =
					XmlUtils.marshaltoW3CDomDocument(o, Context.jc, W, localName, declaredType);
			org.w3c.dom.Element root = doc.getDocumentElement();
			trimNamespaceDeclarations(root);
			return serialize(root);
		} catch (Exception e) {
			note("marshal " + localName + ": " + e);
			return null;
		}
	}

	private static final String XMLNS_NS = "http://www.w3.org/2000/xmlns/";

	/** Remove from the root every xmlns declaration whose prefix nothing under it uses. */
	private static void trimNamespaceDeclarations(org.w3c.dom.Element root) {
		Set<String> used = new TreeSet<String>();
		usedPrefixes(root, used);
		org.w3c.dom.NamedNodeMap attributes = root.getAttributes();
		List<org.w3c.dom.Attr> remove = new ArrayList<org.w3c.dom.Attr>();
		for (int i = 0; i < attributes.getLength(); i++) {
			org.w3c.dom.Attr a = (org.w3c.dom.Attr) attributes.item(i);
			if (!XMLNS_NS.equals(a.getNamespaceURI())) continue;
			String declares = "xmlns".equals(a.getName()) ? "" : a.getLocalName();
			if (!used.contains(declares)) remove.add(a);
		}
		for (org.w3c.dom.Attr a : remove) root.removeAttributeNode(a);
	}

	private static void usedPrefixes(org.w3c.dom.Node node, Set<String> into) {
		if (node.getNodeType() == org.w3c.dom.Node.ELEMENT_NODE
				|| node.getNodeType() == org.w3c.dom.Node.ATTRIBUTE_NODE) {
			String prefix = node.getPrefix();
			if (!XMLNS_NS.equals(node.getNamespaceURI())) into.add(prefix == null ? "" : prefix);
		}
		org.w3c.dom.NamedNodeMap attributes = node.getAttributes();
		if (attributes != null) {
			for (int i = 0; i < attributes.getLength(); i++) usedPrefixes(attributes.item(i), into);
		}
		for (org.w3c.dom.Node c = node.getFirstChild(); c != null; c = c.getNextSibling()) {
			usedPrefixes(c, into);
		}
	}

	/** The element as XML: no declaration, no added white space. */
	private static String serialize(org.w3c.dom.Element element) throws Exception {
		javax.xml.transform.Transformer t =
				javax.xml.transform.TransformerFactory.newInstance().newTransformer();
		t.setOutputProperty(javax.xml.transform.OutputKeys.OMIT_XML_DECLARATION, "yes");
		t.setOutputProperty(javax.xml.transform.OutputKeys.INDENT, "no");
		java.io.StringWriter writer = new java.io.StringWriter();
		t.transform(new javax.xml.transform.dom.DOMSource(element),
				new javax.xml.transform.stream.StreamResult(writer));
		return writer.toString();
	}

	private static String preview(String text) {
		if (text == null) return null;
		return text.length() <= TEXT_PREVIEW ? text : text.substring(0, TEXT_PREVIEW);
	}

	/** A paragraph's or a run's text: every w:t and w:delText under it, a text box's
	 *  content excluded (those paragraphs carry their own). */
	private static String textOf(Object o) {
		StringBuilder sb = new StringBuilder();
		text(o, sb, 0, o instanceof P);
		return sb.toString();
	}

	private static void text(Object o, StringBuilder sb, int depth, boolean skipNestedParagraphs) {
		if (depth > 60) return;
		List<Object> children = childrenOf(o);
		if (children == null) return;
		for (Object raw : children) {
			Object child = XmlUtils.unwrap(raw);
			if (child == null) continue;
			if (child instanceof Text) {
				String value = ((Text) child).getValue();
				if (value != null) sb.append(value);
			} else if (child instanceof org.docx4j.wml.DelText) {
				// w:delText is its own class, not a Text: a deleted run's text is still text
				String value = ((org.docx4j.wml.DelText) child).getValue();
				if (value != null) sb.append(value);
			} else if (isTextBox(child) || (skipNestedParagraphs && child instanceof P)) {
				continue;
			} else {
				text(child, sb, depth + 1, true);
			}
		}
	}

}
