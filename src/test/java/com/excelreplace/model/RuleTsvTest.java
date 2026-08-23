package com.excelreplace.model;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RuleTsvTest {

    @Test
    void roundTripsRules() {
        ReplaceRule regex = new ReplaceRule("ABC-(\\d+)", "XYZ-$1", true);
        ReplaceRule literal = new ReplaceRule("a.b", "x", false);
        literal.setEnabled(false);
        String tsv = RuleTsv.format(List.of(regex, literal));
        List<ReplaceRule> parsed = RuleTsv.parse(tsv);
        assertEquals(2, parsed.size());
        assertTrue(parsed.get(0).isEnabled());
        assertTrue(parsed.get(0).isRegex());
        assertEquals("ABC-(\\d+)", parsed.get(0).getPatternText());
        assertEquals("XYZ-$1", parsed.get(0).getReplacement());
        assertFalse(parsed.get(1).isEnabled());
        assertFalse(parsed.get(1).isRegex());
        assertEquals("a.b", parsed.get(1).getPatternText());
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
        List<ReplaceRule> parsed = RuleTsv.parse("TRUE\tFALSE\t\"a\tb\"\t\"c\td\"\n");
        assertEquals(1, parsed.size());
        assertEquals("a\tb", parsed.get(0).getPatternText());
        assertEquals("c\td", parsed.get(0).getReplacement());
        assertFalse(parsed.get(0).isRegex());
    }
}
