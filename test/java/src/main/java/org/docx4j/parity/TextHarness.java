/*
 * core-ts CR-012 phase 6 (its D8 c): the docx4j oracle for the text model, Texts.of and
 * PlainText.of (docx4j CR-046's org.docx4j.model.text, released in docx4j 17.3.4).
 *
 * For every .docx, .docm and .dotx under each of the fixtures directories, at any depth, records
 * what docx4j's Texts.of answers in four readings (READINGS: the three views, and one with every
 * story and every optional item included), each as the block model - every block's address, story,
 * kind, paragraph fields, runs, fields, numbering and table rows, the comments and the warnings -
 * and as PlainText.of writes it, and the result's metadata once.  test/text-parity.test.mjs reads
 * the same documents through Texts.of and PlainText.of and compares.
 *
 *   mvn -q compile exec:java -Dexec.mainClass=org.docx4j.parity.TextHarness \
 *       -Dfixtures=../fixtures,../text-extraction -Dout=../golden/text -Ddocx4j.commit=<hash>
 *
 * A golden's path is the fixtures directory's name, then the document's path under it, its
 * extension becoming .json: ../fixtures/parity/numbering-stories.docx is
 * fixtures/parity/numbering-stories.json.
 *
 * A value is written only where it is not its default (null, false, an empty list, and 0 or 1
 * where the class says that is the default), so that a golden holds what a block says rather than
 * what it doesn't; the test writes its side by the same rule.  A block's getText() is left out (it
 * is its runs' text, or its cells'), as is TextBlock's section, which has no getter, a chart's
 * ChartDescription and an image's source object.  A reading whose result is the same as an earlier
 * reading's, to the character, is written as {"options": ..., "sameAs": "<that reading>"}: most
 * documents have no revisions, so their three views agree.  Json.writeFolded puts each small object
 * on one line.
 */
package org.docx4j.parity;

import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.TimeZone;

import org.docx4j.Docx4jProperties;
import org.docx4j.jaxb.McSelection;
import org.docx4j.model.text.ImageDescription;
import org.docx4j.model.text.PlainText;
import org.docx4j.model.text.TextBlock;
import org.docx4j.model.text.TextCell;
import org.docx4j.model.text.TextComment;
import org.docx4j.model.text.TextField;
import org.docx4j.model.text.TextImage;
import org.docx4j.model.text.TextNumbering;
import org.docx4j.model.text.TextOptions;
import org.docx4j.model.text.TextResult;
import org.docx4j.model.text.TextRevision;
import org.docx4j.model.text.TextRow;
import org.docx4j.model.text.TextRun;
import org.docx4j.model.text.TextWarning;
import org.docx4j.model.text.Texts;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;

public class TextHarness {

	/** Raise when the shape of a golden changes; the test reads it. */
	static final String HARNESS_VERSION = "1";

	static final String TIME_ZONE = "UTC";

	/**
	 * The readings, by name: each an option set in the WD's names and values (REQ-1001), which
	 * the golden records and the test passes to this package's Texts.of as they stand.
	 * {@code everything} is the default view with every story included and every item that is
	 * off by default turned on, but for {@code embedded} (REQ-512) and {@code fields: update}
	 * (REQ-506), which change what is read rather than what is shown.
	 */
	static final Map<String, Map<String, Object>> READINGS = new LinkedHashMap<String, Map<String, Object>>();
	static {
		READINGS.put("accepted", new LinkedHashMap<String, Object>());
		Map<String, Object> original = new LinkedHashMap<String, Object>();
		original.put("view", "original");
		READINGS.put("original", original);
		Map<String, Object> markup = new LinkedHashMap<String, Object>();
		markup.put("view", "markup");
		markup.put("markup_detail", "author_date");
		READINGS.put("markup", markup);
		Map<String, Object> everything = new LinkedHashMap<String, Object>();
		everything.put("include", List.of("headers", "footers", "comments", "glossary"));
		everything.put("comments", true);
		everything.put("hidden", true);
		everything.put("field_codes", true);
		everything.put("toc", true);
		everything.put("list_labels", true);
		everything.put("alt_text", true);
		everything.put("metadata", true);
		everything.put("page_markers", true);
		READINGS.put("everything", everything);
	}

	public static void main(String[] args) throws Exception {
		TimeZone.setDefault(TimeZone.getTimeZone(TIME_ZONE));
		Docx4jProperties.setProperty("docx4j.fonts.discoverPhysicalFonts.enabled", "false");
		Docx4jProperties.setProperty("docx4j.fonts.discoverJarFonts.enabled", "true");
		System.setProperty("docx4j.fonts.fontcache", Files.createTempDirectory("docx4j-parity-fontcache").toString());
		Docx4jProperties.setProperty(McSelection.PROPERTY, Harness.MC_PREFER_CHOICE);

		File out = new File(Harness.required("out"));
		String commit = System.getProperty("docx4j.commit", "unknown");
		Files.createDirectories(out.toPath());

		List<String[]> documents = new ArrayList<String[]>();   // { golden path, fixture path, file }
		for (String dir : Harness.required("fixtures").split(",")) {
			File root = new File(dir.trim());
			if (!root.isDirectory()) throw new IllegalArgumentException("not a directory: " + root);
			String rootName = root.getCanonicalFile().getName();
			List<File> files = new ArrayList<File>();
			documents(root, files);
			for (File f : files) {
				String rel = rootName + "/" + root.toPath().relativize(f.toPath()).toString().replace(File.separatorChar, '/');
				documents.add(new String[] { rel.substring(0, rel.lastIndexOf('.')) + ".json", rel, f.getPath() });
			}
		}
		documents.sort(Comparator.comparing(d -> d[0]));
		for (int i = 1; i < documents.size(); i++) {
			if (documents.get(i)[0].equals(documents.get(i - 1)[0])) {
				throw new IllegalStateException("two documents, one golden: " + documents.get(i - 1)[1] + ", " + documents.get(i)[1]);
			}
		}
		System.out.println("docx4j commit " + commit + "; " + documents.size() + " documents -> " + out);

		long started = System.currentTimeMillis();
		int failed = 0;
		for (String[] d : documents) {
			long t = System.currentTimeMillis();
			try {
				Map<String, Object> golden = golden(new File(d[2]), d[1], commit);
				Path file = out.toPath().resolve(d[0]);
				Files.createDirectories(file.getParent());
				Files.write(file, Json.writeFolded(golden).getBytes(StandardCharsets.UTF_8));
				System.out.println(String.format("  %-80s %5d ms  %8d bytes", d[1], System.currentTimeMillis() - t, Files.size(file)));
			} catch (Exception e) {
				failed++;
				System.out.println("  " + d[1] + " FAILED: " + e);
				e.printStackTrace();
			}
		}
		System.out.println((documents.size() - failed) + " goldens in "
				+ (System.currentTimeMillis() - started) + " ms" + (failed == 0 ? "" : ", " + failed + " FAILED"));
		if (failed > 0) System.exit(1);
	}

	/** Every WordprocessingML document under dir, at any depth. */
	private static void documents(File dir, List<File> into) {
		File[] files = dir.listFiles();
		if (files == null) return;
		for (File f : files) {
			String n = f.getName().toLowerCase(Locale.ROOT);
			if (f.isDirectory()) documents(f, into);
			else if (n.endsWith(".docx") || n.endsWith(".docm") || n.endsWith(".dotx")) into.add(f);
		}
	}

	// ----------------------------------------------------------------- per document

	static Map<String, Object> golden(File file, String name, String commit) throws Exception {
		WordprocessingMLPackage pkg = WordprocessingMLPackage.load(file);

		Map<String, Object> readings = new LinkedHashMap<String, Object>();
		Map<String, String> bodies = new LinkedHashMap<String, String>();   // a reading's body as JSON -> its name
		Map<String, Object> metadata = null;
		for (Map.Entry<String, Map<String, Object>> reading : READINGS.entrySet()) {
			TextResult result = Texts.of(pkg, TextOptions.of(reading.getValue()));
			if (metadata == null) metadata = result.metadata();
			Map<String, Object> body = new LinkedHashMap<String, Object>();
			body.put("plainText", PlainText.of(result));
			body.put("blocks", blocks(result.blocks()));
			List<Object> comments = new ArrayList<Object>();
			for (TextComment c : result.comments().values()) comments.add(comment(c));
			put(body, "comments", comments);
			List<Object> warnings = new ArrayList<Object>();
			for (TextWarning w : result.warnings()) warnings.add(warning(w));
			put(body, "warnings", warnings);
			Map<String, Object> r = new LinkedHashMap<String, Object>();
			r.put("options", reading.getValue());
			String json = Json.write(body);
			if (bodies.containsKey(json)) {
				r.put("sameAs", bodies.get(json));
			} else {
				bodies.put(json, reading.getKey());
				r.putAll(body);
			}
			readings.put(reading.getKey(), r);
		}

		Map<String, Object> header = new LinkedHashMap<String, Object>();
		header.put("harnessVersion", HARNESS_VERSION);
		header.put("docx4jCommit", commit);
		header.put("docx4jVersion", Harness.docx4jVersion());
		header.put("docx4jCoreJarSha256", Harness.docx4jCoreJarSha256());
		header.put("date", Instant.now().truncatedTo(ChronoUnit.SECONDS).toString());
		header.put("fixture", name);
		header.put("fixtureBytes", Files.size(file.toPath()));
		header.put("timeZone", TIME_ZONE);
		header.put("mcPreferChoice", Harness.MC_PREFER_CHOICE);

		Map<String, Object> golden = new LinkedHashMap<String, Object>();
		golden.put("header", header);
		golden.put("metadata", value(metadata));
		golden.put("readings", readings);
		return golden;
	}

	// ----------------------------------------------------------------- the block model

	private static List<Object> blocks(List<TextBlock> blocks) {
		List<Object> out = new ArrayList<Object>();
		for (TextBlock b : blocks) out.add(block(b));
		return out;
	}

	private static Map<String, Object> block(TextBlock b) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("address", b.getAddress());
		m.put("kind", b.getKind().name());
		m.put("story", b.getStory().name());
		put(m, "paraId", b.getParaId());
		put(m, "style", b.getStyle());
		put(m, "hidden", b.isHidden());
		put(m, "placeholder", b.isPlaceholder());
		put(m, "decorative", b.isDecorative());
		if (b.getRevisions() != 0) m.put("revisions", b.getRevisions());
		put(m, "indentLeft", b.getIndentLeft());
		put(m, "indentHanging", b.getIndentHanging());
		put(m, "indentFirstLine", b.getIndentFirstLine());
		put(m, "alignment", b.getAlignment());
		put(m, "monospace", b.isMonospace());
		put(m, "listLabel", b.getListLabel());
		put(m, "outlineLevel", b.getOutlineLevel());
		if (b.getNumbering() != null) m.put("numbering", numbering(b.getNumbering()));
		put(m, "contextualSpacing", b.getContextualSpacing());
		put(m, "bottomBorder", b.hasBottomBorder());
		put(m, "noteType", b.getNoteType());
		if (b.getNoteNumber() != 0) m.put("noteNumber", b.getNoteNumber());
		put(m, "noteId", b.getNoteId());
		put(m, "simpleForm", b.isSimpleForm());
		if (b.getKind() == TextBlock.Kind.CHART) {
			put(m, "chartTitle", b.getChartTitle());
			put(m, "chartHeader", new ArrayList<Object>(b.getChartHeader()));
			List<Object> rows = new ArrayList<Object>();
			for (List<String> row : b.getChartRows()) rows.add(new ArrayList<Object>(row));
			put(m, "chartRows", rows);
		}
		List<Object> runs = new ArrayList<Object>();
		for (TextRun r : b.getRuns()) runs.add(run(r));
		put(m, "runs", runs);
		List<Object> fields = new ArrayList<Object>();
		for (TextField f : b.getFields()) fields.add(field(f));
		put(m, "fields", fields);
		List<Object> rows = new ArrayList<Object>();
		for (TextRow row : b.getRows()) rows.add(row(row));
		put(m, "rows", rows);
		return m;
	}

	private static Map<String, Object> run(TextRun r) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("kind", r.getKind().name());
		put(m, "text", r.getText());
		if (r.getRevision() != null) m.put("revision", revision(r.getRevision()));
		put(m, "hidden", r.isHidden());
		put(m, "placeholder", r.isPlaceholder());
		put(m, "noteType", r.getNoteType());
		if (r.getNoteNumber() != 0) m.put("noteNumber", r.getNoteNumber());
		put(m, "noteId", r.getNoteId());
		put(m, "commentId", r.getCommentId());
		if (r.getImage() != null) m.put("image", image(r.getImage()));
		if (r.getDescription() != null) m.put("description", description(r.getDescription()));
		put(m, "progId", r.getProgId());
		if (r.getPageNumber() != 0) m.put("pageNumber", r.getPageNumber());
		put(m, "bold", r.isBold());
		put(m, "italic", r.isItalic());
		put(m, "strike", r.isStrike());
		put(m, "monospace", r.isMonospace());
		put(m, "link", r.getLink());
		put(m, "math", r.isMath());
		put(m, "displayMath", r.isDisplayMath());
		return m;
	}

	private static Map<String, Object> revision(TextRevision v) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("type", v.getType().name());
		put(m, "author", v.getAuthor());
		put(m, "date", v.getDate());
		return m;
	}

	private static Map<String, Object> image(TextImage i) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		put(m, "descr", i.getDescr());
		put(m, "name", i.getName());
		put(m, "title", i.getTitle());
		put(m, "decorative", i.isDecorative());
		put(m, "vml", i.isVml());
		if (i.getPart() != null) m.put("part", i.getPart().getPartName().getName());
		return m;
	}

	private static Map<String, Object> description(ImageDescription d) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("text", d.getText());
		m.put("provenance", d.getProvenance().name());
		return m;
	}

	private static Map<String, Object> field(TextField f) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("instr", f.getInstr());
		put(m, "result", f.getResult());
		put(m, "updated", f.isUpdated());
		return m;
	}

	private static Map<String, Object> numbering(TextNumbering n) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("numId", n.getNumId());
		m.put("level", n.getLevel());
		put(m, "format", n.getFormat());
		put(m, "levelText", n.getLevelText());
		m.put("start", n.getStart());
		put(m, "value", n.getValue());
		put(m, "abstractNumId", n.getAbstractNumId());
		put(m, "bullet", n.isBullet());
		put(m, "label", n.isLabel());
		return m;
	}

	private static Map<String, Object> row(TextRow row) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("address", row.getAddress());
		put(m, "header", row.isHeader());
		List<Object> cells = new ArrayList<Object>();
		for (TextCell c : row.getCells()) {
			Map<String, Object> cell = new LinkedHashMap<String, Object>();
			cell.put("address", c.getAddress());
			if (c.getGridSpan() != 1) cell.put("gridSpan", c.getGridSpan());
			put(cell, "mergeContinue", c.isMergeContinue());
			put(cell, "mergeRestart", c.isMergeRestart());
			put(cell, "blocks", blocks(c.getBlocks()));
			cells.add(cell);
		}
		put(m, "cells", cells);
		return m;
	}

	private static Map<String, Object> comment(TextComment c) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("id", c.getId());
		put(m, "author", c.getAuthor());
		put(m, "initials", c.getInitials());
		put(m, "date", c.getDate());
		put(m, "parentId", c.getParentId());
		put(m, "resolved", c.isResolved());
		put(m, "blocks", blocks(c.getBlocks()));
		List<Object> replies = new ArrayList<Object>();
		for (TextComment reply : c.getReplies()) replies.add(reply.getId());
		put(m, "replies", replies);
		return m;
	}

	private static Map<String, Object> warning(TextWarning w) {
		Map<String, Object> m = new LinkedHashMap<String, Object>();
		m.put("kind", w.getKind());
		put(m, "address", w.getAddress());
		put(m, "branch", w.getBranch());
		put(m, "contentType", w.getContentType());
		put(m, "partName", w.getPartName());
		return m;
	}

	/** The metadata as JSON: maps in their order, numbers and booleans as they are, anything else as its string. */
	private static Object value(Object v) {
		if (v == null || v instanceof String || v instanceof Boolean) return v;
		if (v instanceof Number) {
			// BigInteger, BigDecimal, Float and Double print as JSON reads them; keep the number
			return v;
		}
		if (v instanceof Map) {
			Map<String, Object> m = new LinkedHashMap<String, Object>();
			for (Map.Entry<?, ?> e : ((Map<?, ?>) v).entrySet()) m.put(String.valueOf(e.getKey()), value(e.getValue()));
			return m;
		}
		if (v instanceof List) {
			List<Object> l = new ArrayList<Object>();
			for (Object o : (List<?>) v) l.add(value(o));
			return l;
		}
		return v.toString();
	}

	/** Puts a value unless it is a default: null, false, or an empty list. */
	private static void put(Map<String, Object> m, String key, Object value) {
		if (value == null) return;
		if (value instanceof Boolean && !((Boolean) value)) return;
		if (value instanceof List && ((List<?>) value).isEmpty()) return;
		m.put(key, value);
	}
}
