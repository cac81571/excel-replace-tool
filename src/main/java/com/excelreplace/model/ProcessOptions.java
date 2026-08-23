package com.excelreplace.model;

import java.awt.Color;
import java.util.regex.Pattern;

public final class ProcessOptions {
    private boolean cells = true;
    private boolean shapes = true;
    private boolean comments = true;
    private boolean headersFooters = true;
    private boolean sheetNames = false;
    private boolean caseInsensitive = false;
    private boolean multiline = true;
    private boolean recolor = true;
    private Color replacementColor = new Color(220, 20, 60);

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

    public boolean isCaseInsensitive() {
        return caseInsensitive;
    }

    public void setCaseInsensitive(boolean caseInsensitive) {
        this.caseInsensitive = caseInsensitive;
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

    public int regexFlags() {
        int flags = Pattern.UNICODE_CHARACTER_CLASS;
        if (caseInsensitive) {
            flags |= Pattern.CASE_INSENSITIVE | Pattern.UNICODE_CASE;
        }
        if (multiline) {
            flags |= Pattern.MULTILINE;
        }
        return flags;
    }
}
