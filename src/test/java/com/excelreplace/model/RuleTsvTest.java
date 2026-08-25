package com.excelreplace.model;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RuleTsvTest {

    @Test
    void roundTripsRules() {
        ReplaceRule regex = new ReplaceRule("ABC-(\\d+)", "XYZ-$1", true, true);
        regex.setTargetSheets(List.of("画面設計"));
        regex.setCellRanges(List.of("A1:C10", "E5"));
        ReplaceRule literal = new ReplaceRule("a.b", "x", false, false);
        literal.setEnabled(false);
        String tsv = RuleTsv.format(List.of(regex, literal));
        List<ReplaceRule> parsed = RuleTsv.parse(tsv);
        assertEquals(2, parsed.size());
        assertTrue(parsed.get(0).isEnabled());
        assertTrue(parsed.get(0).isRegex());
        assertTrue(parsed.get(0).isIgnoreCase());
        assertEquals("ABC-(\\d+)", parsed.get(0).getPatternText());
        assertEquals("XYZ-$1", parsed.get(0).getReplacement());
        assertEquals(List.of("画面設計"), parsed.get(0).getTargetSheets());
        assertEquals(List.of("A1:C10", "E5"), parsed.get(0).getCellRanges());
        assertFalse(parsed.get(1).isEnabled());
        assertFalse(parsed.get(1).isRegex());
        assertFalse(parsed.get(1).isIgnoreCase());
        assertEquals("a.b", parsed.get(1).getPatternText());
        assertTrue(parsed.get(1).getTargetSheets().isEmpty());
    }

    @Test
    void parsesTwoColumnExcelPaste() {
        List<ReplaceRule> parsed = RuleTsv.parse("旧システム\t新システム\n帳票\t伝票\n");
        assertEquals(2, parsed.size());
        assertEquals("旧システム", parsed.get(0).getPatternText());
        assertEquals("新システム", parsed.get(0).getReplacement());
        assertTrue(parsed.get(0).isRegex());
    }

    @Test
    void parsesQuotedFieldWithTab() {
        List<ReplaceRule> parsed = RuleTsv.parse("TRUE\tFALSE\tTRUE\t\"a\tb\"\t\"c\td\"\n");
        assertEquals(1, parsed.size());
        assertEquals("a\tb", parsed.get(0).getPatternText());
        assertEquals("c\td", parsed.get(0).getReplacement());
        assertFalse(parsed.get(0).isRegex());
        assertTrue(parsed.get(0).isIgnoreCase());
    }

    @Test
    void parsesLegacyFourColumnRows() {
        List<ReplaceRule> parsed = RuleTsv.parse("TRUE\tFALSE\ta.b\tx\n");
        assertEquals(1, parsed.size());
        assertFalse(parsed.get(0).isRegex());
        assertFalse(parsed.get(0).isIgnoreCase());
        assertEquals("a.b", parsed.get(0).getPatternText());
    }
}
