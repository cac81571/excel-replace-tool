package com.excelreplace.excel;

import com.excelreplace.model.ReplaceRule;
import org.junit.jupiter.api.Test;

import java.awt.Color;
import java.util.List;
import java.util.regex.Pattern;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RichTextEngineTest {

    @Test
    void replacesWithGroupAndKeepsUnmatchedText() {
        List<RichTextEngine.Run> input = RichTextEngine.singleRun("ID: ABC-001 / ABC-002", null);
        ReplaceRule rule = new ReplaceRule("ABC-(\\d+)", "XYZ-$1");
        RichTextEngine.ApplyResult result = RichTextEngine.apply(
                input, RichTextEngine.compile(List.of(rule), 0, true, Color.RED));
        assertEquals("ID: XYZ-001 / XYZ-002", result.text());
        assertEquals(2, result.replacementCount());
        assertTrue(result.runs().stream().anyMatch(run -> run.replaced && Color.RED.equals(run.replacementColor)));
    }

    @Test
    void appliesRulesInOrderWithSharedColor() {
        List<RichTextEngine.Run> input = RichTextEngine.singleRun("旧名称", null);
        ReplaceRule first = new ReplaceRule("旧名称", "仮名称");
        ReplaceRule second = new ReplaceRule("仮名称", "新名称");
        RichTextEngine.ApplyResult result = RichTextEngine.apply(
                input, RichTextEngine.compile(List.of(first, second), 0, true, Color.BLUE));
        assertEquals("新名称", result.text());
        assertEquals(Color.BLUE, result.runs().get(0).replacementColor);
    }

    @Test
    void expandsBackslashN() {
        RichTextEngine.ApplyResult result = RichTextEngine.applyRule(
                RichTextEngine.singleRun("a b", null),
                Pattern.compile(" "),
                "\\n",
                false,
                null);
        assertEquals("a\nb", result.text());
    }

    @Test
    void treatsLiteralPatternAsPlainText() {
        ReplaceRule regex = new ReplaceRule("a.b", "X", true);
        ReplaceRule literal = new ReplaceRule("a.b", "Y", false);
        assertEquals("X X", RichTextEngine.apply(
                RichTextEngine.singleRun("a.b axb", null),
                RichTextEngine.compile(List.of(regex), 0, false, null)).text());
        assertEquals("Y axb", RichTextEngine.apply(
                RichTextEngine.singleRun("a.b axb", null),
                RichTextEngine.compile(List.of(literal), 0, false, null)).text());
    }

    @Test
    void doesNotExpandGroupsWhenRegexOff() {
        ReplaceRule rule = new ReplaceRule("hello", "$1", false);
        assertEquals("$1", RichTextEngine.apply(
                RichTextEngine.singleRun("hello", null),
                RichTextEngine.compile(List.of(rule), 0, false, null)).text());
    }

    @Test
    void skipsDisabledAndBlankRules() {
        ReplaceRule disabled = new ReplaceRule("A", "B");
        disabled.setEnabled(false);
        ReplaceRule blank = new ReplaceRule("   ", "C");
        assertTrue(RichTextEngine.compile(List.of(disabled, blank), 0).isEmpty());
    }

    @Test
    void ignoresCaseOnlyWhenRuleRequestsIt() {
        ReplaceRule sensitive = new ReplaceRule("abc", "X", false, false);
        ReplaceRule insensitive = new ReplaceRule("abc", "Y", false, true);
        assertEquals("X Abc", RichTextEngine.apply(
                RichTextEngine.singleRun("abc Abc", null),
                RichTextEngine.compile(List.of(sensitive), 0, false, null)).text());
        assertEquals("Y Y", RichTextEngine.apply(
                RichTextEngine.singleRun("abc Abc", null),
                RichTextEngine.compile(List.of(insensitive), 0, false, null)).text());
    }

    @Test
    void compiledRuleMatchesSheetAndCellScope() {
        ReplaceRule rule = new ReplaceRule("a", "b", false);
        rule.setTargetSheets(List.of("画面設計"));
        rule.setCellRanges(List.of("B2:C3"));
        RichTextEngine.CompiledRule compiled = RichTextEngine.compile(List.of(rule), 0, false, null).get(0);
        assertTrue(compiled.matchesSheet("画面設計"));
        assertTrue(!compiled.matchesSheet("帳票一覧"));
        assertTrue(compiled.matchesCell(1, 1));
        assertTrue(!compiled.matchesCell(0, 0));
        assertTrue(!compiled.appliesOutsideCells());
        assertTrue(compiled.appliesToSheet("画面設計", true));
        assertTrue(!compiled.appliesToSheet("帳票一覧", true));
    }

    @Test
    void allSheetRuleIsBlockedByGlobalExcludeButExplicitRuleIsNot() {
        ReplaceRule all = new ReplaceRule("a", "b", false);
        ReplaceRule explicit = new ReplaceRule("a", "c", false);
        explicit.setTargetSheets(List.of("改訂履歴"));
        List<RichTextEngine.CompiledRule> compiled = RichTextEngine.compile(List.of(all, explicit), 0, false, null);
        List<RichTextEngine.CompiledRule> onExcluded = RichTextEngine.filterForSheet(compiled, "改訂履歴", true);
        assertEquals(1, onExcluded.size());
        assertEquals("c", onExcluded.get(0).replacement);
    }
}
