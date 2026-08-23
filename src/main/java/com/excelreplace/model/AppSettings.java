package com.excelreplace.model;

import java.awt.Color;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

/**
 * 置換ルール・対象オプション・文字色などの設定ファイル。
 */
public final class AppSettings {
    static final String HEADER = "# excel-replace-tool settings";
    static final String RULES_MARK = "[rules]";

    private boolean recursive;
    private boolean dumpText = true;
    private String inputPath = "";
    private String outputPath = "";
    private final ProcessOptions options = new ProcessOptions();
    private final List<ReplaceRule> rules = new ArrayList<>();

    public boolean isRecursive() {
        return recursive;
    }

    public void setRecursive(boolean recursive) {
        this.recursive = recursive;
    }

    public boolean isDumpText() {
        return dumpText;
    }

    public void setDumpText(boolean dumpText) {
        this.dumpText = dumpText;
    }

    public String getInputPath() {
        return inputPath;
    }

    public void setInputPath(String inputPath) {
        this.inputPath = inputPath == null ? "" : inputPath;
    }

    public String getOutputPath() {
        return outputPath;
    }

    public void setOutputPath(String outputPath) {
        this.outputPath = outputPath == null ? "" : outputPath;
    }

    public ProcessOptions getOptions() {
        return options;
    }

    public List<ReplaceRule> getRules() {
        return rules;
    }

    public String format() {
        StringBuilder out = new StringBuilder();
        out.append(HEADER).append('\n');
        out.append("version=1\n");
        write(out, "recursive", recursive);
        write(out, "dumpText", dumpText);
        write(out, "inputPath", inputPath);
        write(out, "outputPath", outputPath);
        write(out, "cells", options.isCells());
        write(out, "shapes", options.isShapes());
        write(out, "comments", options.isComments());
        write(out, "headersFooters", options.isHeadersFooters());
        write(out, "sheetNames", options.isSheetNames());
        write(out, "caseInsensitive", options.isCaseInsensitive());
        write(out, "multiline", options.isMultiline());
        write(out, "recolor", options.isRecolor());
        out.append("replacementColor=").append(toHex(options.getReplacementColor())).append('\n');
        out.append('\n').append(RULES_MARK).append('\n');
        out.append(RuleTsv.format(rules));
        return out.toString();
    }

    public static AppSettings parse(String text) {
        AppSettings settings = new AppSettings();
        if (text == null || text.isBlank()) {
            return settings;
        }
        String normalized = text.replace("\r\n", "\n").replace('\r', '\n');
        int mark = indexOfRulesMark(normalized);
        String optionsPart = mark < 0 ? normalized : normalized.substring(0, mark);
        String rulesPart = mark < 0 ? "" : normalized.substring(mark + RULES_MARK.length()).stripLeading();
        Map<String, String> values = readKeyValues(optionsPart);
        settings.setRecursive(bool(values.get("recursive"), false));
        settings.setDumpText(bool(values.get("dumpText"), true));
        settings.setInputPath(values.getOrDefault("inputPath", ""));
        settings.setOutputPath(values.getOrDefault("outputPath", ""));
        ProcessOptions options = settings.getOptions();
        options.setCells(bool(values.get("cells"), true));
        options.setShapes(bool(values.get("shapes"), true));
        options.setComments(bool(values.get("comments"), true));
        options.setHeadersFooters(bool(values.get("headersFooters"), true));
        options.setSheetNames(bool(values.get("sheetNames"), false));
        options.setCaseInsensitive(bool(values.get("caseInsensitive"), false));
        options.setMultiline(bool(values.get("multiline"), true));
        options.setRecolor(bool(values.get("recolor"), true));
        options.setReplacementColor(parseColor(values.get("replacementColor"), options.getReplacementColor()));
        if (rulesPart.isBlank() && !looksLikeSettingsFile(normalized)) {
            settings.getRules().addAll(RuleTsv.parse(normalized));
        } else {
            settings.getRules().addAll(RuleTsv.parse(rulesPart));
        }
        return settings;
    }

    private static boolean looksLikeSettingsFile(String text) {
        String first = text.stripLeading();
        return first.startsWith("#") || first.startsWith("version=") || first.contains(RULES_MARK);
    }

    private static int indexOfRulesMark(String text) {
        int index = 0;
        while (index < text.length()) {
            int lineEnd = text.indexOf('\n', index);
            if (lineEnd < 0) {
                lineEnd = text.length();
            }
            if (text.substring(index, lineEnd).trim().equals(RULES_MARK)) {
                return index;
            }
            index = lineEnd + 1;
        }
        return -1;
    }

    private static Map<String, String> readKeyValues(String text) {
        Map<String, String> values = new LinkedHashMap<>();
        for (String line : text.split("\n", -1)) {
            String trimmed = line.trim();
            if (trimmed.isEmpty() || trimmed.startsWith("#") || trimmed.equals(RULES_MARK)) {
                continue;
            }
            int eq = trimmed.indexOf('=');
            if (eq <= 0) {
                continue;
            }
            values.put(trimmed.substring(0, eq).trim(), trimmed.substring(eq + 1).trim());
        }
        return values;
    }

    private static void write(StringBuilder out, String key, boolean value) {
        out.append(key).append('=').append(value).append('\n');
    }

    private static void write(StringBuilder out, String key, String value) {
        out.append(key).append('=').append(value == null ? "" : value).append('\n');
    }

    private static boolean bool(String value, boolean defaultValue) {
        if (value == null || value.isBlank()) {
            return defaultValue;
        }
        return RuleTsv.parseBoolean(value, defaultValue);
    }

    static String toHex(Color color) {
        Color value = color == null ? Color.RED : color;
        return String.format("#%06X", value.getRGB() & 0xFFFFFF);
    }

    static Color parseColor(String token, Color defaultValue) {
        if (token == null || token.isBlank()) {
            return defaultValue;
        }
        String value = token.trim();
        try {
            if (value.startsWith("#") && value.length() == 7) {
                return Color.decode(value);
            }
            if (value.matches("(?i)[0-9a-f]{6}")) {
                return Color.decode("#" + value);
            }
        } catch (RuntimeException ignored) {
            // fall through
        }
        return defaultValue;
    }
}
