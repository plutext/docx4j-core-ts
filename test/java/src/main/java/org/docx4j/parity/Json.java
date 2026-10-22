/*
 * CR-001 Phase B step 1: the parity harness's JSON writer.
 *
 * Hand written so that the harness depends on docx4j alone (one offline build, no
 * serialisation library on the classpath to version-drift).  Values are the plain Java
 * shapes the harness builds: LinkedHashMap (an object, written in insertion order, which
 * is what makes a golden stable and readable), List (an array), String, Number, Boolean
 * and null.
 */
package org.docx4j.parity;

import java.util.List;
import java.util.Map;

final class Json {

	private Json() {}

	/** Pretty printed, two spaces an indent, a trailing newline: a file to read in a diff. */
	static String write(Object value) {
		StringBuilder out = new StringBuilder(1 << 16);
		write(value, out, 0);
		out.append('\n');
		return out.toString();
	}

	/** The width {@link #writeFolded} folds to. */
	static final int FOLD = 120;

	/**
	 * Pretty printed as {@link #write} is, but an object or array whose one-line form fits in
	 * {@link #FOLD} columns at its indent is written on that line: TextHarness's goldens, whose runs
	 * and numberings are many small objects, read as one line each and are a third the size.
	 */
	static String writeFolded(Object value) {
		StringBuilder out = new StringBuilder(1 << 16);
		folded(value, out, 0);
		out.append('\n');
		return out.toString();
	}

	private static void folded(Object value, StringBuilder out, int depth) {
		boolean container = (value instanceof Map && !((Map<?, ?>) value).isEmpty())
				|| (value instanceof List && !((List<?>) value).isEmpty());
		if (!container) {
			write(value, out, depth);
			return;
		}
		StringBuilder line = new StringBuilder();
		oneLine(value, line);
		if (line.length() + 2 * depth <= FOLD) {
			out.append(line);
			return;
		}
		boolean map = value instanceof Map;
		out.append(map ? "{\n" : "[\n");
		boolean first = true;
		for (Object o : map ? ((Map<?, ?>) value).entrySet() : (List<?>) value) {
			if (!first) out.append(",\n");
			first = false;
			indent(out, depth + 1);
			if (map) {
				Map.Entry<?, ?> e = (Map.Entry<?, ?>) o;
				string(String.valueOf(e.getKey()), out);
				out.append(": ");
				folded(e.getValue(), out, depth + 1);
			} else {
				folded(o, out, depth + 1);
			}
		}
		out.append('\n');
		indent(out, depth);
		out.append(map ? '}' : ']');
	}

	private static void oneLine(Object value, StringBuilder out) {
		if (value instanceof Map) {
			out.append('{');
			boolean first = true;
			for (Map.Entry<?, ?> e : ((Map<?, ?>) value).entrySet()) {
				if (!first) out.append(", ");
				first = false;
				string(String.valueOf(e.getKey()), out);
				out.append(": ");
				oneLine(e.getValue(), out);
			}
			out.append('}');
		} else if (value instanceof List) {
			out.append('[');
			boolean first = true;
			for (Object o : (List<?>) value) {
				if (!first) out.append(", ");
				first = false;
				oneLine(o, out);
			}
			out.append(']');
		} else {
			write(value, out, 0);
		}
	}

	private static void write(Object value, StringBuilder out, int depth) {

		if (value == null) {
			out.append("null");
		} else if (value instanceof String) {
			string((String) value, out);
		} else if (value instanceof Boolean || value instanceof Number) {
			out.append(value.toString());
		} else if (value instanceof Map) {
			Map<?, ?> map = (Map<?, ?>) value;
			if (map.isEmpty()) {
				out.append("{}");
				return;
			}
			out.append("{\n");
			boolean first = true;
			for (Map.Entry<?, ?> e : map.entrySet()) {
				if (!first) out.append(",\n");
				first = false;
				indent(out, depth + 1);
				string(String.valueOf(e.getKey()), out);
				out.append(": ");
				write(e.getValue(), out, depth + 1);
			}
			out.append('\n');
			indent(out, depth);
			out.append('}');
		} else if (value instanceof List) {
			List<?> list = (List<?>) value;
			if (list.isEmpty()) {
				out.append("[]");
				return;
			}
			out.append("[\n");
			boolean first = true;
			for (Object o : list) {
				if (!first) out.append(",\n");
				first = false;
				indent(out, depth + 1);
				write(o, out, depth + 1);
			}
			out.append('\n');
			indent(out, depth);
			out.append(']');
		} else {
			// never reached by the harness; a defensive last resort rather than a silent wrong type
			string(value.toString(), out);
		}
	}

	private static void indent(StringBuilder out, int depth) {
		for (int i = 0; i < depth; i++) out.append("  ");
	}

	/** JSON string, with every non-ASCII and every control character escaped, so the
	 *  goldens are pure ASCII and diff the same everywhere. */
	private static void string(String s, StringBuilder out) {
		out.append('"');
		for (int i = 0; i < s.length(); i++) {
			char c = s.charAt(i);
			switch (c) {
				case '"':  out.append("\\\""); break;
				case '\\': out.append("\\\\"); break;
				case '\b': out.append("\\b"); break;
				case '\f': out.append("\\f"); break;
				case '\n': out.append("\\n"); break;
				case '\r': out.append("\\r"); break;
				case '\t': out.append("\\t"); break;
				default:
					if (c < 0x20 || c > 0x7e) {
						out.append(String.format("\\u%04x", (int) c));
					} else {
						out.append(c);
					}
			}
		}
		out.append('"');
	}
}
