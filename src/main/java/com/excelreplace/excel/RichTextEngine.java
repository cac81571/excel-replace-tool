package com.excelreplace.excel;

import com.excelreplace.model.ProcessOptions;
import com.excelreplace.model.ReplaceRule;
import org.apache.poi.ss.util.CellRangeAddress;

import java.awt.Color;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 書式付きテキストに対して正規表現置換を行い、置換箇所の色情報を保持する。
 */
public final class RichTextEngine {

    private RichTextEngine() {
    }

    public static final class Run {
        public final String text;
        public final Object fontKey;
        public final boolean replaced;
        public final Color replacementColor;

        public Run(String text, Object fontKey, boolean replaced, Color replacementColor) {
            this.text = text == null ? "" : text;
            this.fontKey = fontKey;
            this.replaced = replaced;
            this.replacementColor = replacementColor;
        }
    }

    public static final class CompiledRule {
        public final Pattern pattern;
        public final String replacement;
        public final boolean recolor;
        public final Color color;
        public final boolean regex;
        private final List<String> targetSheets;
        private final List<CellRangeAddress> cellRanges;

        public CompiledRule(
                Pattern pattern,
                String replacement,
                boolean recolor,
                Color color,
                boolean regex,
                List<String> targetSheets,
                List<CellRangeAddress> cellRanges) {
            this.pattern = pattern;
            this.replacement = replacement;
            this.recolor = recolor;
            this.color = color;
            this.regex = regex;
            this.targetSheets = targetSheets == null ? List.of() : List.copyOf(targetSheets);
            this.cellRanges = cellRanges == null ? List.of() : List.copyOf(cellRanges);
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

        /** 対象シートを明示指定しているルールかどうか。 */
        public boolean hasExplicitSheets() {
            return !targetSheets.isEmpty();
        }

        /**
         * シートへの適用可否。全体の「置換対象外」より、ルールの対象シート指定を優先する。
         * 対象外シートでも、ルールがシート名を明示していれば適用する。
         */
        public boolean appliesToSheet(String sheetName, boolean sheetGloballyExcluded) {
            if (!matchesSheet(sheetName)) {
                return false;
            }
            return !sheetGloballyExcluded || hasExplicitSheets();
        }

        public boolean hasCellRangeLimit() {
            return !cellRanges.isEmpty();
        }

        public boolean matchesCell(int rowIndex, int columnIndex) {
            if (cellRanges.isEmpty()) {
                return true;
            }
            for (CellRangeAddress range : cellRanges) {
                if (range.isInRange(rowIndex, columnIndex)) {
                    return true;
                }
            }
            return false;
        }

        /** セル範囲指定があるルールはセル（とそのセルのコメント）以外には適用しない。 */
        public boolean appliesOutsideCells() {
            return cellRanges.isEmpty();
        }
    }

    public record ApplyResult(List<Run> runs, int replacementCount) {
        public String text() {
            return concat(runs);
        }

        public boolean changed() {
            return replacementCount > 0;
        }
    }

    public static List<Run> singleRun(String text, Object fontKey) {
        List<Run> runs = new ArrayList<>();
        if (text != null && !text.isEmpty()) {
            runs.add(new Run(text, fontKey, false, null));
        }
        return runs;
    }

    public static String concat(List<Run> runs) {
        StringBuilder sb = new StringBuilder();
        for (Run run : runs) {
            sb.append(run.text);
        }
        return sb.toString();
    }

    public static List<CompiledRule> compile(List<ReplaceRule> rules, int flags) {
        return compile(rules, flags, false, null);
    }

    public static List<CompiledRule> compile(List<ReplaceRule> rules, ProcessOptions options) {
        return compile(rules, options.regexFlags(), options.isRecolor(), options.getReplacementColor());
    }

    public static List<CompiledRule> compile(
            List<ReplaceRule> rules, int flags, boolean recolor, Color color) {
        List<CompiledRule> compiled = new ArrayList<>();
        for (int i = 0; i < rules.size(); i++) {
            ReplaceRule rule = rules.get(i);
            if (!rule.isEnabled() || rule.getPatternText().isBlank()) {
                continue;
            }
            try {
                compiled.add(new CompiledRule(
                        rule.compile(flags),
                        rule.getReplacement(),
                        recolor,
                        color,
                        rule.isRegex(),
                        rule.getTargetSheets(),
                        parseCellRanges(rule.getCellRanges(), i + 1)));
            } catch (IllegalArgumentException e) {
                throw e;
            } catch (Exception e) {
                throw new IllegalArgumentException(
                        "ルール " + (i + 1) + " の正規表現が不正です: " + rule.getPatternText(), e);
            }
        }
        return compiled;
    }

    private static List<CellRangeAddress> parseCellRanges(List<String> ranges, int ruleNumber) {
        List<CellRangeAddress> parsed = new ArrayList<>();
        if (ranges == null) {
            return parsed;
        }
        for (String range : ranges) {
            if (range == null || range.isBlank()) {
                continue;
            }
            try {
                parsed.add(CellRangeAddress.valueOf(range.trim().toUpperCase(Locale.ROOT)));
            } catch (RuntimeException e) {
                throw new IllegalArgumentException(
                        "ルール " + ruleNumber + " のセル範囲が不正です: " + range, e);
            }
        }
        return parsed;
    }

    public static List<CompiledRule> filterForSheet(
            List<CompiledRule> rules, String sheetName, boolean sheetGloballyExcluded) {
        List<CompiledRule> filtered = new ArrayList<>();
        for (CompiledRule rule : rules) {
            if (rule.appliesToSheet(sheetName, sheetGloballyExcluded)) {
                filtered.add(rule);
            }
        }
        return filtered;
    }

    public static List<CompiledRule> filterForCell(List<CompiledRule> rules, int rowIndex, int columnIndex) {
        List<CompiledRule> filtered = new ArrayList<>();
        for (CompiledRule rule : rules) {
            if (rule.matchesCell(rowIndex, columnIndex)) {
                filtered.add(rule);
            }
        }
        return filtered;
    }

    public static List<CompiledRule> filterOutsideCells(List<CompiledRule> rules) {
        List<CompiledRule> filtered = new ArrayList<>();
        for (CompiledRule rule : rules) {
            if (rule.appliesOutsideCells()) {
                filtered.add(rule);
            }
        }
        return filtered;
    }

    public static ApplyResult apply(List<Run> input, List<CompiledRule> rules) {
        List<Run> current = copy(input);
        int total = 0;
        for (CompiledRule rule : rules) {
            ApplyResult result = applyRule(
                    current,
                    rule.pattern,
                    rule.replacement,
                    rule.recolor,
                    rule.color,
                    rule.regex);
            current = result.runs();
            total += result.replacementCount();
        }
        return new ApplyResult(current, total);
    }

    static ApplyResult applyRule(
            List<Run> input,
            Pattern pattern,
            String replacementTemplate,
            boolean recolor,
            Color color) {
        return applyRule(input, pattern, replacementTemplate, recolor, color, true);
    }

    static ApplyResult applyRule(
            List<Run> input,
            Pattern pattern,
            String replacementTemplate,
            boolean recolor,
            Color color,
            boolean regex) {
        String text = concat(input);
        if (text.isEmpty()) {
            return new ApplyResult(input, 0);
        }

        Run[] runAt = mapCharsToRuns(input, text.length());
        Matcher matcher = pattern.matcher(text);
        StringBuilder out = new StringBuilder(text.length());
        List<StyledRange> ranges = new ArrayList<>();
        int last = 0;
        int count = 0;

        while (matcher.find()) {
            if (matcher.start() == matcher.end()) {
                continue;
            }
            appendUnmatched(out, ranges, text, last, matcher.start(), runAt);
            String replacement = regex
                    ? expandReplacement(matcher, replacementTemplate)
                    : (replacementTemplate == null ? "" : replacementTemplate);
            int start = out.length();
            out.append(replacement);
            int end = out.length();
            if (end > start) {
                Object fontKey = runAt[matcher.start()] == null ? null : runAt[matcher.start()].fontKey;
                ranges.add(new StyledRange(start, end, fontKey, true, recolor ? color : null));
            }
            last = matcher.end();
            count++;
        }
        appendUnmatched(out, ranges, text, last, text.length(), runAt);
        return new ApplyResult(mergeRanges(out.toString(), ranges), count);
    }

    static String expandReplacement(Matcher matcher, String template) {
        if (template == null || template.isEmpty()) {
            return "";
        }
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < template.length(); i++) {
            char c = template.charAt(i);
            if (c == '\\') {
                if (i + 1 >= template.length()) {
                    sb.append('\\');
                    break;
                }
                char next = template.charAt(++i);
                switch (next) {
                    case 'n' -> sb.append('\n');
                    case 'r' -> sb.append('\r');
                    case 't' -> sb.append('\t');
                    default -> sb.append(next);
                }
            } else if (c == '$') {
                if (i + 1 < template.length() && template.charAt(i + 1) == '{') {
                    int end = template.indexOf('}', i + 2);
                    if (end < 0) {
                        throw new IllegalArgumentException("閉じられていない名前付きグループ: " + template);
                    }
                    String name = template.substring(i + 2, end);
                    String group = matcher.group(name);
                    if (group != null) {
                        sb.append(group);
                    }
                    i = end;
                } else {
                    int j = i + 1;
                    if (j >= template.length() || !Character.isDigit(template.charAt(j))) {
                        sb.append('$');
                        continue;
                    }
                    int groupIndex = 0;
                    while (j < template.length() && Character.isDigit(template.charAt(j))) {
                        int next = groupIndex * 10 + (template.charAt(j) - '0');
                        if (next > matcher.groupCount()) {
                            break;
                        }
                        groupIndex = next;
                        j++;
                    }
                    String group = matcher.group(groupIndex);
                    if (group != null) {
                        sb.append(group);
                    }
                    i = j - 1;
                }
            } else {
                sb.append(c);
            }
        }
        return sb.toString();
    }

    private static List<Run> copy(List<Run> input) {
        return new ArrayList<>(input);
    }

    private static Run[] mapCharsToRuns(List<Run> runs, int length) {
        Run[] map = new Run[length];
        int i = 0;
        for (Run run : runs) {
            for (int k = 0; k < run.text.length(); k++) {
                map[i++] = run;
            }
        }
        return map;
    }

    private static void appendUnmatched(
            StringBuilder out,
            List<StyledRange> ranges,
            String text,
            int from,
            int to,
            Run[] runAt) {
        int i = from;
        while (i < to) {
            Run run = runAt[i];
            int j = i + 1;
            while (j < to && runAt[j] == run) {
                j++;
            }
            int start = out.length();
            out.append(text, i, j);
            ranges.add(new StyledRange(
                    start,
                    out.length(),
                    run == null ? null : run.fontKey,
                    run != null && run.replaced,
                    run == null ? null : run.replacementColor));
            i = j;
        }
    }

    private static List<Run> mergeRanges(String text, List<StyledRange> ranges) {
        List<Run> runs = new ArrayList<>();
        StyledRange pending = null;
        for (StyledRange range : ranges) {
            if (range.start == range.end) {
                continue;
            }
            if (pending != null && pending.end == range.start && pending.sameStyle(range)) {
                pending = new StyledRange(pending.start, range.end, pending.fontKey, pending.replaced, pending.color);
            } else {
                if (pending != null) {
                    runs.add(pending.toRun(text));
                }
                pending = range;
            }
        }
        if (pending != null) {
            runs.add(pending.toRun(text));
        }
        return runs;
    }

    private record StyledRange(int start, int end, Object fontKey, boolean replaced, Color color) {
        boolean sameStyle(StyledRange other) {
            return replaced == other.replaced
                    && Objects.equals(fontKey, other.fontKey)
                    && Objects.equals(color, other.color);
        }

        Run toRun(String text) {
            return new Run(text.substring(start, end), fontKey, replaced, color);
        }
    }
}
