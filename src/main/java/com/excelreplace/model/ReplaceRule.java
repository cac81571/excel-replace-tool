package com.excelreplace.model;

import java.util.regex.Pattern;

public final class ReplaceRule {
    private boolean enabled = true;
    private boolean regex = true;
    private String patternText = "";
    private String replacement = "";

    public ReplaceRule() {
    }

    public ReplaceRule(String patternText, String replacement) {
        this(patternText, replacement, true);
    }

    public ReplaceRule(String patternText, String replacement, boolean regex) {
        this.patternText = patternText == null ? "" : patternText;
        this.replacement = replacement == null ? "" : replacement;
        this.regex = regex;
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

    public Pattern compile(int flags) {
        String source = regex ? patternText : Pattern.quote(patternText);
        return Pattern.compile(source, flags);
    }

    public ReplaceRule copy() {
        ReplaceRule copy = new ReplaceRule(patternText, replacement, regex);
        copy.setEnabled(enabled);
        return copy;
    }
}
