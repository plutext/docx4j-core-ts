/*
 * core-ts CR-011 section 4: the docx4j oracle for getAllTrackedChanges and the tracked-change
 * filters (docx4j CR-043's org.docx4j.model.revisions, released in docx4j 17.3.3).
 *
 * For every .docx under the fixtures directory, at any depth, whose word/ parts hold revision
 * markup, records what docx4j's TrackedChanges.in lists - each change's type, kind, author, date,
 * id, text, part and the kinds of its pieces - and then, each on a fresh load of the fixture,
 * what is left listed after
 *
 *   acceptAll and rejectAll, with no filter;
 *   rejectAll(authors(the second author listed)), where the listing has two or more;
 *   acceptAll(types(DELETED));
 *   acceptAll(kinds(RUN_DELETION)), which splits a deleted paragraph: its text goes, its mark
 *     stays (CR-043's U19),
 *
 * with the count each returned, and the comments each leaves (version 2, core-ts CR-002 section 54:
 * docx4j CR-043 section 15's rule): the comment parts the main document part still has, the ids of
 * the comments part's entries, and the ids of the comment range starts, range ends and references
 * left in the package's XML parts under word/.  test/revisions-parity.test.mjs does the same through
 * getAllTrackedChanges, acceptAllRevisions and rejectAllRevisions and compares.
 *
 *   mvn -q compile exec:java -Dexec.mainClass=org.docx4j.parity.RevisionsHarness \
 *       -Dfixtures=../fixtures -Dout=../golden/revisions -Ddocx4j.commit=<hash>
 *
 * A golden's path mirrors the fixture's under the fixtures directory, .docx becoming .json.
 *
 * Dates: TrackedChange.getDate() reads a w:date without w16du:dateUtc as local wall-clock time
 * in the JVM's time zone, as Office JS reads it in the add-in's.  The harness sets the default
 * time zone to UTC before anything reads one, and records it in the header; the test sets TZ=UTC
 * for the same reason.
 */
package org.docx4j.parity;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.Instant;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TimeZone;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

import org.docx4j.Docx4jProperties;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.McSelection;
import org.docx4j.model.revisions.TrackedChange;
import org.docx4j.model.revisions.TrackedChangeFilter;
import org.docx4j.model.revisions.TrackedChanges;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.JaxbXmlPart;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.wml.Comments;

public class RevisionsHarness {

	/** Raise when the shape of a golden changes; the test reads it. */
	static final String HARNESS_VERSION = "2";

	static final String TIME_ZONE = "UTC";

	/**
	 * Revision markup, by local name in any prefix: what makes a fixture's golden worth having.
	 * The same names as core-ts's REVISION_MARKUP (WordprocessingMLPackage.mts), read here over
	 * every XML part under word/.
	 */
	static final Pattern REVISION_MARKUP = Pattern.compile(
			"<(?:[\\w.-]+:)?(?:ins|del|moveFrom|moveTo|rPrChange|pPrChange|sectPrChange|tblPrChange|trPrChange|tcPrChange"
			+ "|tblGridChange|tblPrExChange|cellIns|cellDel|cellMerge|numberingChange|customXmlInsRangeStart|customXmlDelRangeStart)[\\s/>]");

	public static void main(String[] args) throws Exception {
		TimeZone.setDefault(TimeZone.getTimeZone(TIME_ZONE));
		Docx4jProperties.setProperty(McSelection.PROPERTY, Harness.MC_PREFER_CHOICE);

		File fixtures = new File(Harness.required("fixtures"));
		File out = new File(Harness.required("out"));
		String commit = System.getProperty("docx4j.commit", "unknown");
		if (!fixtures.isDirectory()) throw new IllegalArgumentException("not a directory: " + fixtures);
		Files.createDirectories(out.toPath());

		List<File> documents = new ArrayList<File>();
		revised(fixtures, documents);
		documents.sort(Comparator.comparing(f -> relative(fixtures, f)));
		System.out.println("docx4j commit " + commit + "; " + documents.size() + " fixtures with revision markup -> " + out);

		long started = System.currentTimeMillis();
		int failed = 0;
		for (File document : documents) {
			long t = System.currentTimeMillis();
			String name = relative(fixtures, document);
			try {
				Map<String, Object> golden = golden(document, name, commit);
				Path file = out.toPath().resolve(name.substring(0, name.length() - ".docx".length()) + ".json");
				Files.createDirectories(file.getParent());
				Files.write(file, Json.write(golden).getBytes(StandardCharsets.UTF_8));
				System.out.println(String.format("  %-60s %5d ms  %4d changes",
						name, System.currentTimeMillis() - t, ((List<?>) golden.get("listing")).size()));
			} catch (Exception e) {
				failed++;
				System.out.println("  " + name + " FAILED: " + e);
				e.printStackTrace();
			}
		}
		System.out.println((documents.size() - failed) + " goldens in "
				+ (System.currentTimeMillis() - started) + " ms" + (failed == 0 ? "" : ", " + failed + " FAILED"));
		if (failed > 0) System.exit(1);
	}

	/** Every .docx under dir, at any depth, holding revision markup in a part under word/. */
	private static void revised(File dir, List<File> into) throws IOException {
		File[] files = dir.listFiles();
		if (files == null) return;
		for (File f : files) {
			if (f.isDirectory()) revised(f, into);
			else if (f.getName().toLowerCase().endsWith(".docx") && holdsRevisions(f)) into.add(f);
		}
	}

	private static boolean holdsRevisions(File docx) throws IOException {
		try (ZipFile zip = new ZipFile(docx)) {
			for (ZipEntry entry : java.util.Collections.list(zip.entries())) {
				String n = entry.getName();
				if (!n.startsWith("word/") || !n.endsWith(".xml")) continue;
				String xml = new String(zip.getInputStream(entry).readAllBytes(), StandardCharsets.UTF_8);
				if (REVISION_MARKUP.matcher(xml).find()) return true;
			}
		}
		return false;
	}

	private static String relative(File root, File f) {
		return root.toPath().relativize(f.toPath()).toString().replace(File.separatorChar, '/');
	}

	// ----------------------------------------------------------------- per document

	static Map<String, Object> golden(File file, String name, String commit) throws Exception {
		List<TrackedChange> listed = TrackedChanges.in(WordprocessingMLPackage.load(file));
		List<Object> listing = changes(listed);

		Map<String, Object> resolved = new LinkedHashMap<String, Object>();
		{
			WordprocessingMLPackage pkg = WordprocessingMLPackage.load(file);
			resolved.put("acceptAll", outcome(TrackedChanges.acceptAll(pkg), pkg, null));
		}
		{
			WordprocessingMLPackage pkg = WordprocessingMLPackage.load(file);
			resolved.put("rejectAll", outcome(TrackedChanges.rejectAll(pkg), pkg, null));
		}
		Set<String> authors = new LinkedHashSet<String>();
		for (TrackedChange change : listed) for (TrackedChange piece : change.getPieces()) authors.add(piece.getAuthor());
		if (authors.size() >= 2) {
			String author = new ArrayList<String>(authors).get(1);
			WordprocessingMLPackage pkg = WordprocessingMLPackage.load(file);
			Map<String, Object> filter = new LinkedHashMap<String, Object>();
			filter.put("authors", List.of(author));
			resolved.put("rejectAuthor", outcome(TrackedChanges.rejectAll(pkg, new TrackedChangeFilter().authors(author)), pkg, filter));
		}
		{
			WordprocessingMLPackage pkg = WordprocessingMLPackage.load(file);
			Map<String, Object> filter = new LinkedHashMap<String, Object>();
			filter.put("types", List.of("DELETED"));
			resolved.put("acceptDeleted", outcome(TrackedChanges.acceptAll(pkg,
					new TrackedChangeFilter().types(TrackedChange.Type.DELETED)), pkg, filter));
		}
		{
			WordprocessingMLPackage pkg = WordprocessingMLPackage.load(file);
			Map<String, Object> filter = new LinkedHashMap<String, Object>();
			filter.put("kinds", List.of("RUN_DELETION"));
			resolved.put("acceptRunDeletions", outcome(TrackedChanges.acceptAll(pkg,
					new TrackedChangeFilter().kinds(TrackedChange.Kind.RUN_DELETION)), pkg, filter));
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
		header.put("authors", new ArrayList<Object>(authors));

		Map<String, Object> golden = new LinkedHashMap<String, Object>();
		golden.put("header", header);
		golden.put("listing", listing);
		golden.put("resolved", resolved);
		return golden;
	}

	/** What a resolve-all returned, with the filter it was given, and what is listed after it. */
	private static Map<String, Object> outcome(int count, WordprocessingMLPackage pkg, Map<String, Object> filter) throws Exception {
		Map<String, Object> out = new LinkedHashMap<String, Object>();
		if (filter != null) out.put("filter", filter);
		out.put("count", count);
		out.put("left", changes(TrackedChanges.in(pkg)));
		out.put("comments", comments(pkg));
		return out;
	}

	/** A comment marker in marshalled XML, in any prefix: its local name and its id. */
	static final Pattern COMMENT_MARKER = Pattern.compile(
			"<(?:[\\w.-]+:)?(commentRangeStart|commentRangeEnd|commentReference)\\b[^>]*?\\b(?:[\\w.-]+:)?id=\"(-?\\d+)\"");

	/**
	 * What a resolution leaves of the comments: the names of the comment parts the main document part
	 * still has (comments, commentsExtended, commentsIds, commentsExtensible, people), the ids of the
	 * comments part's entries, and the ids of the range starts, range ends and references in the
	 * package's XML parts under word/; each list sorted, an id once.
	 */
	static Map<String, Object> comments(WordprocessingMLPackage pkg) throws Exception {
		MainDocumentPart main = pkg.getMainDocumentPart();
		List<Object> parts = new ArrayList<Object>();
		for (Part part : new Part[] { main.getCommentsPart(), main.getCommentsExtendedPart(), main.getCommentsIdsPart(),
				main.getCommentsExtensiblePart(), main.getPeoplePart() }) {
			if (part != null) parts.add(part.getPartName().getName());
		}
		Set<Long> entries = new TreeSet<Long>();
		if (main.getCommentsPart() != null && main.getCommentsPart().getJaxbElement() != null) {
			for (Comments.Comment comment : main.getCommentsPart().getJaxbElement().getComment()) {
				if (comment.getId() != null) entries.add(comment.getId().longValue());
			}
		}
		Map<String, Set<Long>> markers = new LinkedHashMap<String, Set<Long>>();
		for (String kind : new String[] { "commentRangeStart", "commentRangeEnd", "commentReference" }) markers.put(kind, new TreeSet<Long>());
		for (Part part : pkg.getParts().getParts().values()) {
			if (!(part instanceof JaxbXmlPart) || !part.getPartName().getName().startsWith("/word/")) continue;
			Object root = ((JaxbXmlPart<?>) part).getJaxbElement();
			if (root == null) continue;
			Matcher m = COMMENT_MARKER.matcher(XmlUtils.marshaltoString(root, true, false, ((JaxbXmlPart<?>) part).getJAXBContext()));
			while (m.find()) markers.get(m.group(1)).add(Long.parseLong(m.group(2)));
		}
		Map<String, Object> out = new LinkedHashMap<String, Object>();
		out.put("parts", parts);
		out.put("entries", new ArrayList<Object>(entries));
		out.put("starts", new ArrayList<Object>(markers.get("commentRangeStart")));
		out.put("ends", new ArrayList<Object>(markers.get("commentRangeEnd")));
		out.put("references", new ArrayList<Object>(markers.get("commentReference")));
		return out;
	}

	private static List<Object> changes(List<TrackedChange> list) {
		List<Object> out = new ArrayList<Object>();
		for (TrackedChange change : list) out.add(change(change));
		return out;
	}

	/** One change, as the test reads it back from getAllTrackedChanges. */
	static Map<String, Object> change(TrackedChange change) {
		Map<String, Object> out = new LinkedHashMap<String, Object>();
		out.put("type", change.getType().name());
		out.put("kind", change.getKind().name());
		out.put("author", change.getAuthor());
		Instant date = change.getDate();
		out.put("date", date == null ? null : date.toString());
		out.put("id", change.getId() == null ? null : change.getId().longValue());
		out.put("text", change.getText());
		out.put("part", change.getPart() == null ? null : change.getPart().getPartName().getName());
		List<Object> pieces = new ArrayList<Object>();
		for (TrackedChange piece : change.getPieces()) pieces.add(piece.getKind().name());
		out.put("pieces", pieces);
		return out;
	}
}
