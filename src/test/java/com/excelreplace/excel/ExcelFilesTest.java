package com.excelreplace.excel;

import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelFilesTest {

    @TempDir
    Path temp;

    @Test
    void listsSelectedFilesAndFolderTogether() throws Exception {
        Path folder = Files.createDirectory(temp.resolve("docs"));
        Path a = Files.writeString(folder.resolve("a.xlsx"), "x");
        Path b = Files.writeString(folder.resolve("b.xls"), "x");
        Files.writeString(folder.resolve("a_replaced.xlsx"), "x");
        Path nested = Files.createDirectory(folder.resolve("sub"));
        Path c = Files.writeString(nested.resolve("c.xlsx"), "x");
        Path extra = Files.writeString(temp.resolve("extra.xlsx"), "x");

        List<Path> fromFolder = ExcelFiles.listExcelFiles(folder, false);
        assertEquals(2, fromFolder.size());
        assertTrue(fromFolder.contains(a.toAbsolutePath().normalize()));
        assertTrue(fromFolder.contains(b.toAbsolutePath().normalize()));
        assertFalse(fromFolder.stream().anyMatch(ExcelFiles::isGeneratedOutput));

        List<Path> recursive = ExcelFiles.listExcelFiles(folder, true);
        assertEquals(3, recursive.size());
        assertTrue(recursive.contains(c.toAbsolutePath().normalize()));

        List<Path> mixed = ExcelFiles.listExcelFiles(List.of(folder, extra), false);
        assertEquals(3, mixed.size());
        assertTrue(mixed.contains(extra.toAbsolutePath().normalize()));
    }

    @Test
    void dumpPathsIncludeExcelExtension() {
        Path input = temp.resolve("sample.xlsx");
        Path output = temp.resolve("out").resolve("sample_replaced.xlsx");
        assertEquals(temp.resolve("out").resolve("sample_xlsx.txt"), ExcelFiles.dumpPathBeside(input, output));
        assertEquals(temp.resolve("out").resolve("sample_replaced_xlsx.txt"), ExcelFiles.defaultDumpPath(output));
        assertEquals("sample_xls.txt", ExcelFiles.dumpFileName("sample.xls"));
        assertEquals("設計書サンプル_xlsm.txt", ExcelFiles.dumpFileName("設計書サンプル.xlsm"));
        assertEquals("sample_xlsx.txt", ExcelFiles.dumpFileName("sample.XLSX"));
    }

    @Test
    void createParentDirectoriesMakesMissingFolders() throws Exception {
        Path file = temp.resolve("missing").resolve("nested").resolve("a.xlsx");
        assertFalse(Files.exists(file.getParent()));
        ExcelFiles.createParentDirectories(file);
        assertTrue(Files.isDirectory(file.getParent()));
        ExcelFiles.createParentDirectories(file);
        assertTrue(Files.isDirectory(file.getParent()));
    }
}
