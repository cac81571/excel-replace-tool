package com.excelreplace.ui;

import com.excelreplace.model.ReplaceRule;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertArrayEquals;

class RuleTableModelTest {

    @Test
    void movesSelectedRowsUpAndDown() {
        RuleTableModel model = new RuleTableModel();
        model.replaceAll(java.util.List.of(
                new ReplaceRule("a", "1"),
                new ReplaceRule("b", "2"),
                new ReplaceRule("c", "3")));

        assertArrayEquals(new int[] {0, 1}, model.moveRowsUp(new int[] {1, 2}));
        assertEquals("b", model.getRules().get(0).getPatternText());
        assertEquals("c", model.getRules().get(1).getPatternText());
        assertEquals("a", model.getRules().get(2).getPatternText());

        assertArrayEquals(new int[] {1, 2}, model.moveRowsDown(new int[] {0, 1}));
        assertEquals("a", model.getRules().get(0).getPatternText());
        assertEquals("b", model.getRules().get(1).getPatternText());
        assertEquals("c", model.getRules().get(2).getPatternText());
    }

    @Test
    void doesNotMovePastEdges() {
        RuleTableModel model = new RuleTableModel();
        model.replaceAll(java.util.List.of(
                new ReplaceRule("a", "1"),
                new ReplaceRule("b", "2")));

        assertEquals(0, model.moveRowsUp(new int[] {0}).length);
        assertEquals("a", model.getRules().get(0).getPatternText());
        assertEquals(0, model.moveRowsDown(new int[] {1}).length);
        assertEquals("b", model.getRules().get(1).getPatternText());
    }
}
