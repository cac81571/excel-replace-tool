package com.excelreplace.model;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;

/**
 * 置換ルールの TSV 入出力。設定ファイルの [rules] 節や、Excel から書き出した TSV を読み込む。
 */
public final class RuleTsv {
    public static final String HEADER = "有効\t正規表現\t検索\t置換後";

    private RuleTsv() {
    }

    public static String format(List<ReplaceRule> rules) {
        StringBuilder out = new StringBuilder();
        out.append(HEADER).append('\n');
        for (ReplaceRule rule : rules) {
            if (rule.getPatternText().isBlank() && rule.getReplacement().isBlank()) {
                continue;
            }
            out.append(escape(bool(rule.isEnabled()))).append('\t')
                    .append(escape(bool(rule.isRegex()))).append('\t')
                    .append(escape(rule.getPatternText())).append('\t')
                    .append(escape(rule.getReplacement()))
                    .append('\n');
        }
        return out.toString();
    }

    public static List<ReplaceRule> parse(String tsv) {
        List<List<String>> rows = readRows(tsv == null ? "" : tsv);
        if (rows.isEmpty()) {
            return List.of();
        }
        int start = 0;
        ColumnMap columns = null;
        if (looksLikeHeader(rows.get(0))) {
            columns = ColumnMap.fromHeader(rows.get(0));
            start = 1;
        }
        List<ReplaceRule> rules = new ArrayList<>();
        for (int i = start; i < rows.size(); i++) {
            ReplaceRule rule = toRule(rows.get(i), columns);
            if (rule != null) {
                rules.add(rule);
            }
        }
        return rules;
    }

    private static ReplaceRule toRule(List<String> row, ColumnMap columns) {
        if (row.isEmpty() || row.stream().allMatch(String::isBlank)) {
            return null;
        }
        String enabled;
        String regex;
        String pattern;
        String replacement;
        if (columns != null) {
            enabled = columns.get(row, "enabled");
            regex = columns.get(row, "regex");
            pattern = columns.get(row, "pattern");
            replacement = columns.get(row, "replacement");
        } else if (row.size() >= 4) {
            enabled = row.get(0);
            regex = row.get(1);
            pattern = row.get(2);
            replacement = row.get(3);
        } else if (row.size() == 3) {
            if (isBooleanToken(row.get(2))) {
                enabled = "TRUE";
                regex = row.get(2);
                pattern = row.get(0);
                replacement = row.get(1);
            } else {
                enabled = row.get(0);
                regex = "TRUE";
                pattern = row.get(1);
                replacement = row.get(2);
            }
        } else if (row.size() == 2) {
            enabled = "TRUE";
            regex = "TRUE";
            pattern = row.get(0);
            replacement = row.get(1);
        } else {
            enabled = "TRUE";
            regex = "TRUE";
            pattern = row.get(0);
            replacement = "";
        }
        if (pattern == null || pattern.isBlank()) {
            return null;
        }
        ReplaceRule rule = new ReplaceRule(pattern, replacement == null ? "" : replacement, parseBoolean(regex, true));
        rule.setEnabled(parseBoolean(enabled, true));
        return rule;
    }

    static boolean looksLikeHeader(List<String> row) {
        for (String cell : row) {
            String key = normalizeHeader(cell);
            if (key.equals("enabled") || key.equals("regex") || key.equals("pattern") || key.equals("replacement")) {
                return true;
            }
        }
        return false;
    }

    static boolean parseBoolean(String token, boolean defaultValue) {
        if (token == null || token.isBlank()) {
            return defaultValue;
        }
        String value = token.trim().toLowerCase(Locale.ROOT);
        return switch (value) {
            case "true", "t", "1", "yes", "y", "on", "はい", "有効", "○", "〇", "正規表現" -> true;
            case "false", "f", "0", "no", "n", "off", "いいえ", "無効", "×", "x", "リテラル" -> false;
            default -> defaultValue;
        };
    }

    static boolean isBooleanToken(String token) {
        if (token == null || token.isBlank()) {
            return false;
        }
        String value = token.trim().toLowerCase(Locale.ROOT);
        return switch (value) {
            case "true", "t", "1", "yes", "y", "on", "はい", "有効", "○", "〇", "正規表現",
                    "false", "f", "0", "no", "n", "off", "いいえ", "無効", "×", "x", "リテラル" -> true;
            default -> false;
        };
    }

    static String normalizeHeader(String header) {
        if (header == null) {
            return "";
        }
        String value = header.trim().toLowerCase(Locale.ROOT).replace(" ", "");
        return switch (value) {
            case "有効", "enabled", "enable", "on" -> "enabled";
            case "正規表現", "regex", "regexp", "re" -> "regex";
            case "検索", "検索文字列", "pattern", "search", "from", "置換前" -> "pattern";
            case "置換後", "置換", "replacement", "replace", "to" -> "replacement";
            default -> value;
        };
    }

    static List<List<String>> readRows(String tsv) {
        List<List<String>> rows = new ArrayList<>();
        String normalized = tsv.replace("\r\n", "\n").replace('\r', '\n');
        List<String> row = new ArrayList<>();
        StringBuilder field = new StringBuilder();
        boolean inQuotes = false;
        for (int i = 0; i < normalized.length(); i++) {
            char c = normalized.charAt(i);
            if (inQuotes) {
                if (c == '"') {
                    if (i + 1 < normalized.length() && normalized.charAt(i + 1) == '"') {
                        field.append('"');
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    field.append(c);
                }
            } else if (c == '"') {
                inQuotes = true;
            } else if (c == '\t') {
                row.add(field.toString());
                field.setLength(0);
            } else if (c == '\n') {
                row.add(field.toString());
                field.setLength(0);
                rows.add(row);
                row = new ArrayList<>();
            } else {
                field.append(c);
            }
        }
        if (inQuotes || field.length() > 0 || !row.isEmpty()) {
            row.add(field.toString());
            rows.add(row);
        }
        return rows;
    }

    static String escape(String value) {
        String text = value == null ? "" : value;
        if (text.indexOf('\t') >= 0 || text.indexOf('\n') >= 0 || text.indexOf('\r') >= 0 || text.indexOf('"') >= 0) {
            return '"' + text.replace("\"", "\"\"") + '"';
        }
        return text;
    }

    private static String bool(boolean value) {
        return value ? "TRUE" : "FALSE";
    }

    private static final class ColumnMap {
        private final Map<String, Integer> indexes = new LinkedHashMap<>();

        static ColumnMap fromHeader(List<String> header) {
            ColumnMap map = new ColumnMap();
            for (int i = 0; i < header.size(); i++) {
                String key = normalizeHeader(header.get(i));
                if (!key.isBlank()) {
                    map.indexes.put(key, i);
                }
            }
            return map;
        }

        String get(List<String> row, String key) {
            Integer index = indexes.get(key);
            if (index == null || index < 0 || index >= row.size()) {
                return "";
            }
            return row.get(index);
        }
    }
}
