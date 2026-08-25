package com.excelreplace.excel;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelDiffDumperTest {

    @Test
    void emitsOnlyChangedRecordsWithBeforeAndAfter() {
        String before = """
                FILE\tsample.xlsx
                SHEETS\t1
                SHEET\t0\t画面設計\tvisible
                画面設計\tCELL\tA1\tSTRING\t旧システム
                画面設計\tCELL\tB1\tSTRING\t変更なし
                画面設計\tCOMMENT\tC1\tauthor\t旧コメント
                画面設計\tSHAPE\t0\t1,1:3,3\t旧図形
                """;
        String after = """
                FILE\tsample.xlsx
                SHEETS\t1
                SHEET\t0\t画面設計\tvisible
                画面設計\tCELL\tA1\tSTRING\t新システム
                画面設計\tCELL\tB1\tSTRING\t変更なし
                画面設計\tCOMMENT\tC1\tauthor\t新コメント
                画面設計\tSHAPE\t0\t1,1:3,3\t新図形
                """;

        String diff = ExcelDiffDumper.diff(before, after, "sample.xlsx");
        assertTrue(diff.startsWith("FILE\tsample.xlsx\nCHANGED\t3\n"));
        assertTrue(diff.contains("画面設計\tCELL\tA1\tSTRING\t旧システム\t新システム\n"));
        assertTrue(diff.contains("画面設計\tCOMMENT\tC1\tauthor\t旧コメント\tauthor\t新コメント\n"));
        assertTrue(diff.contains("画面設計\tSHAPE\t0\t1,1:3,3\t旧図形\t新図形\n"));
        assertFalse(diff.contains("変更なし"));
    }

    @Test
    void comparesSheetRenameByIndexWithoutDuplicatingUnchangedCells() {
        String before = """
                FILE\tsample.xlsx
                SHEETS\t1
                SHEET\t0\t旧シート\tvisible
                旧シート\tCELL\tA1\tSTRING\t同じ値
                """;
        String after = """
                FILE\tsample.xlsx
                SHEETS\t1
                SHEET\t0\t新シート\tvisible
                新シート\tCELL\tA1\tSTRING\t同じ値
                """;

        String diff = ExcelDiffDumper.diff(before, after, "sample.xlsx");
        assertEquals("""
                FILE\tsample.xlsx
                CHANGED\t1
                SHEET\t0\t旧シート\tvisible\t新シート\tvisible
                """, diff);
    }
}
