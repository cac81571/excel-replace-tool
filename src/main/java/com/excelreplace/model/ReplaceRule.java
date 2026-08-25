package com.excelreplace.model;

import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.regex.Pattern;

public final class ReplaceRule {
    private boolean enabled = true;
    private boolean regex = true;
    private boolean ignoreCase = false;
    private String patternText = "";
    private String replacement = "";
    private final List<String> targetSheets = new ArrayList<>();
    private final List<String> cellRanges = new ArrayList<>();

    public ReplaceRule() {
    }

    public ReplaceRule(String patternText, String replacement) {
        this(patternText, replacement, true);
    }

    public ReplaceRule(String patternText, String replacement, boolean regex) {
        this(patternText, replacement, regex, false);
    }

    public ReplaceRule(String patternText, String replacement, boolean regex, boolean ignoreCase) {
        this.patternText = patternText == null ? "" : patternText;
        this.replacement = replacement == null ? "" : replacement;
        this.regex = regex;
        this.ignoreCase = ignoreCase;
    }

    public boolean isEnabled() {
        return enabled;
    }

    public void setEnabled(boolean enabled) {
        this.enabled = enabled;
    }

    public boolean isRegex() {
        return regex;
    }

    public void setRegex(boolean regex) {
        this.regex = regex;
    }

    public boolean isIgnoreCase() {
        return ignoreCase;
    }

    public void setIgnoreCase(boolean ignoreCase) {
        this.ignoreCase = ignoreCase;
    }

    public String getPatternText() {
        return patternText;
    }

    public void setPatternText(String patternText) {
        this.patternText = patternText == null ? "" : patternText;
    }

    public String getReplacement() {
        return replacement;
    }

    public void setReplacement(String replacement) {
        this.replacement = replacement == null ? "" : replacement;
    }

    public List<String> getTargetSheets() {
        return targetSheets;
    }

    public void setTargetSheets(Collection<String> sheets) {
        targetSheets.clear();
        if (sheets == null) {
            return;
        }
        for (String sheet : sheets) {
            if (sheet != null && !sheet.isBlank()) {
                targetSheets.add(sheet.trim());
            }
        }
    }

    public List<String> getCellRanges() {
        return cellRanges;
    }

    public void setCellRanges(Collection<String> ranges) {
        cellRanges.clear();
        if (ranges == null) {
            return;
        }
        for (String range : ranges) {
            if (range != null && !range.isBlank()) {
                cellRanges.add(range.trim());
            }
        }
    }

    public boolean matchesSheet(String sheetName) {
        if (targetSheets.isEmpty()) {
            return true;
        }
        if (sheetName == null) {
            return false;
        }
        String target = sheetName.trim();
        for (String sheet : targetSheets) {
            if (sheet.equalsIgnoreCase(target)) {
                return true;
            }
        }
        return false;
    }

    public boolean hasCellRangeLimit() {
        return !cellRanges.isEmpty();
    }

    public Pattern compile(int baseFlags) {
        String source = regex ? patternText : Pattern.quote(patternText);
        int flags = baseFlags;
        if (ignoreCase) {
            flags |= Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE;
        }
        return Pattern.compile(source, flags);
    }

    public ReplaceRule copy() {
        ReplaceRule copy = new ReplaceRule(patternText, replacement, regex, ignoreCase);
        copy.setEnabled(enabled);
        copy.setTargetSheets(targetSheets);
        copy.setCellRanges(cellRanges);
        return copy;
    }

    /** セル範囲の入力文字列を分解（; または , 区切り）。空なら全セル。 */
    public static List<String> parseRangeList(String text) {
        List<String> ranges = new ArrayList<>();
        if (text == null || text.isBlank()) {
            return ranges;
        }
        for (String part : text.split("[;；,]")) {
            String value = part.trim();
            if (!value.isEmpty()) {
                ranges.add(value);
            }
        }
        return ranges;
    }

    public static String formatRangeList(Collection<String> ranges) {
        if (ranges == null || ranges.isEmpty()) {
            return "";
        }
        StringBuilder out = new StringBuilder();
        for (String range : ranges) {
            if (range == null || range.isBlank()) {
                continue;
            }
            if (out.length() > 0) {
                out.append("; ");
            }
            out.append(range.trim());
        }
        return out.toString();
    }
}
