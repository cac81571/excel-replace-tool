package com.excelreplace.model;

import org.junit.jupiter.api.Test;

import java.awt.Color;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class AppSettingsTest {

    @Test
    void roundTripsOptionsAndRules() {
        AppSettings original = new AppSettings();
        original.setRecursive(true);
        original.setDumpText(false);
        original.setInputPath("C:\\work\\a.xlsx; C:\\work\\b.xlsx");
        original.setOutputPath("C:\\work\\out");
        original.getOptions().setSheetNames(true);
        original.getOptions().setCaseInsensitive(true);
        original.getOptions().setRecolor(true);
        original.getOptions().setReplacementColor(new Color(0, 128, 255));
        original.getOptions().setExcludedSheets(List.of("改訂履歴", "表紙"));
        ReplaceRule rule = new ReplaceRule("旧システム", "新システム", false);
        original.getRules().add(rule);

        AppSettings parsed = AppSettings.parse(original.format());
        assertTrue(parsed.isRecursive());
        assertFalse(parsed.isDumpText());
        assertEquals("C:\\work\\a.xlsx; C:\\work\\b.xlsx", parsed.getInputPath());
        assertEquals("C:\\work\\out", parsed.getOutputPath());
        assertTrue(parsed.getOptions().isSheetNames());
        assertTrue(parsed.getOptions().isCaseInsensitive());
        assertEquals(new Color(0, 128, 255), parsed.getOptions().getReplacementColor());
        assertEquals(List.of("改訂履歴", "表紙"), parsed.getOptions().getExcludedSheets());
        assertEquals(1, parsed.getRules().size());
        assertEquals("旧システム", parsed.getRules().get(0).getPatternText());
        assertEquals("新システム", parsed.getRules().get(0).getReplacement());
        assertFalse(parsed.getRules().get(0).isRegex());
    }

    @Test
    void importsPlainTsvAsRulesOnly() {
        AppSettings parsed = AppSettings.parse("検索\t置換後\n帳票\t伝票\n");
        assertEquals(1, parsed.getRules().size());
        assertEquals("帳票", parsed.getRules().get(0).getPatternText());
        assertEquals("伝票", parsed.getRules().get(0).getReplacement());
        assertTrue(parsed.getOptions().isCells());
    }
}
