package com.excelreplace.model;

import java.awt.Color;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;
import java.util.regex.Pattern;

public final class ProcessOptions {
    private boolean cells = true;
    private boolean shapes = true;
    private boolean comments = true;
    private boolean headersFooters = true;
    private boolean sheetNames = false;
    private boolean multiline = true;
    private boolean recolor = true;
    private Color replacementColor = new Color(220, 20, 60);
    private final List<String> excludedSheets = new ArrayList<>();

    public boolean isCells() {
        return cells;
    }

    public void setCells(boolean cells) {
        this.cells = cells;
    }

    public boolean isShapes() {
        return shapes;
    }

    public void setShapes(boolean shapes) {
        this.shapes = shapes;
    }

    public boolean isComments() {
        return comments;
    }

    public void setComments(boolean comments) {
        this.comments = comments;
    }

    public boolean isHeadersFooters() {
        return headersFooters;
    }

    public void setHeadersFooters(boolean headersFooters) {
        this.headersFooters = headersFooters;
    }

    public boolean isSheetNames() {
        return sheetNames;
    }

    public void setSheetNames(boolean sheetNames) {
        this.sheetNames = sheetNames;
    }

    public boolean isMultiline() {
        return multiline;
    }

    public void setMultiline(boolean multiline) {
        this.multiline = multiline;
    }

    public boolean isRecolor() {
        return recolor;
    }

    public void setRecolor(boolean recolor) {
        this.recolor = recolor;
    }

    public Color getReplacementColor() {
        return replacementColor;
    }

    public void setReplacementColor(Color replacementColor) {
        this.replacementColor = replacementColor == null ? Color.RED : replacementColor;
    }

    public List<String> getExcludedSheets() {
        return excludedSheets;
    }

    public void setExcludedSheets(Collection<String> sheets) {
        excludedSheets.clear();
        if (sheets == null) {
            return;
        }
        for (String sheet : sheets) {
            if (sheet != null && !sheet.isBlank()) {
                excludedSheets.add(sheet.trim());
            }
        }
    }

    public boolean isSheetExcluded(String sheetName) {
        if (sheetName == null || excludedSheets.isEmpty()) {
            return false;
        }
        String target = sheetName.trim();
        for (String excluded : excludedSheets) {
            if (excluded.equalsIgnoreCase(target)) {
                return true;
            }
        }
        return false;
    }

    public static List<String> parseSheetList(String text) {
        List<String> sheets = new ArrayList<>();
        if (text == null || text.isBlank()) {
            return sheets;
        }
        for (String part : text.split("[;；]")) {
            String name = part.trim();
            if (!name.isEmpty()) {
                sheets.add(name);
            }
        }
        return sheets;
    }

    public static String formatSheetList(Collection<String> sheets) {
        if (sheets == null || sheets.isEmpty()) {
            return "";
        }
        StringBuilder out = new StringBuilder();
        for (String sheet : sheets) {
            if (sheet == null || sheet.isBlank()) {
                continue;
            }
            if (out.length() > 0) {
                out.append("; ");
            }
            out.append(sheet.trim());
        }
        return out.toString();
    }

    public int regexFlags() {
        int flags = Pattern.UNICODE_CHARACTER_CLASS;
        if (multiline) {
            flags |= Pattern.MULTILINE;
        }
        return flags;
    }
}
