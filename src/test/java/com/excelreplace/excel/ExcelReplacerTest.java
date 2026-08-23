package com.excelreplace.excel;

import com.excelreplace.model.ProcessOptions;
import com.excelreplace.model.ProcessResult;
import com.excelreplace.model.ReplaceRule;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelReplacerTest {

    @Test
    void replacesCellsShapesCommentsAndDumpsXlsx() throws Exception {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("画面設計");
        Row row = sheet.createRow(0);
        Cell cell = row.createCell(0);
        cell.setCellValue("旧システム 画面");

        XSSFDrawing drawing = sheet.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 1, 3, 3);
        XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.setShapeType(ShapeTypes.ROUND_RECT);
        shape.setText("旧システム\nログイン");

        addComment(sheet, 0, 1, "旧システム コメント");

        ProcessResult result = replace(workbook);
        assertTrue(result.getCellHits() >= 1);
        assertTrue(result.getShapeHits() >= 1);
        assertTrue(result.getCommentHits() >= 1);
        assertEquals("新システム 画面", sheet.getRow(0).getCell(0).getStringCellValue());
        assertTrue(sheet.getRow(0).getCell(1).getCellComment().getString().getString().contains("新システム"));

        byte[] bytes = roundTripBytes(workbook);
        XSSFWorkbook reopened = (XSSFWorkbook) org.apache.poi.ss.usermodel.WorkbookFactory.create(
                new ByteArrayInputStream(bytes));
        XSSFSheet reopenedSheet = reopened.getSheetAt(0);
        assertEquals("新システム 画面", reopenedSheet.getRow(0).getCell(0).getStringCellValue());
        XSSFSimpleShape reopenedShape = findFirstSimpleShape(reopenedSheet);
        assertEquals("新システム\nログイン", reopenedShape.getText().trim());
        assertDrawingRgbIsSixHexDigits(bytes);

        String dump = new ExcelTextDumper().dump(reopened, "sample.xlsx");
        assertTrue(dump.contains("CELL\tA1\tSTRING\t新システム 画面"));
        assertTrue(dump.contains("SHAPE\t"));
        assertTrue(dump.contains("COMMENT\tB1\t"));
        assertTrue(dump.contains("新システム"));
        workbook.close();
        reopened.close();
    }

    @Test
    void replacesXlsCellsAndShapes() throws Exception {
        HSSFWorkbook workbook = new HSSFWorkbook();
        HSSFSheet sheet = workbook.createSheet("帳票");
        Row row = sheet.createRow(0);
        row.createCell(0).setCellValue("旧システム");

        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        var anchor = patriarch.createAnchor(0, 0, 0, 0, 1, 1, 3, 3);
        HSSFSimpleShape shape = patriarch.createSimpleShape(anchor);
        shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
        shape.setString(new org.apache.poi.hssf.usermodel.HSSFRichTextString("旧システム"));

        ProcessResult result = replace(workbook);
        assertTrue(result.getCellHits() >= 1);
        assertEquals("新システム", sheet.getRow(0).getCell(0).getStringCellValue());
        assertTrue(shape.getString().getString().contains("新システム"));

        String dump = new ExcelTextDumper().dump(roundTrip(workbook), "sample.xls");
        assertTrue(dump.contains("新システム"));
        workbook.close();
    }

    @Test
    void dumpEscapesNewlinesForDiff() {
        assertEquals("a\\nb\\tc", ExcelTextDumper.esc("a\nb\tc"));
    }

    private static ProcessResult replace(Workbook workbook) {
        ReplaceRule rule = new ReplaceRule("旧システム", "新システム");
        ProcessOptions options = new ProcessOptions();
        return new ExcelReplacer().processWorkbook(workbook, List.of(rule), options);
    }

    private static void addComment(Sheet sheet, int rowIndex, int colIndex, String text) {
        CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(colIndex);
        anchor.setRow1(rowIndex);
        anchor.setCol2(colIndex + 2);
        anchor.setRow2(rowIndex + 2);
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(helper.createRichTextString(text));
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellComment(comment);
    }

    private static XSSFSimpleShape findFirstSimpleShape(XSSFSheet sheet) {
        XSSFDrawing drawing = sheet.getDrawingPatriarch();
        for (var shape : drawing.getShapes()) {
            if (shape instanceof XSSFSimpleShape simple && simple.getText() != null && !simple.getText().isBlank()) {
                return simple;
            }
        }
        throw new AssertionError("図形が見つかりません");
    }

    private static Workbook roundTrip(Workbook workbook) throws Exception {
        return org.apache.poi.ss.usermodel.WorkbookFactory.create(new ByteArrayInputStream(roundTripBytes(workbook)));
    }

    private static byte[] roundTripBytes(Workbook workbook) throws Exception {
        ByteArrayOutputStream buffer = new ByteArrayOutputStream();
        workbook.write(buffer);
        return buffer.toByteArray();
    }

    private static void assertDrawingRgbIsSixHexDigits(byte[] xlsx) throws Exception {
        boolean foundDrawing = false;
        try (java.util.zip.ZipInputStream zip = new java.util.zip.ZipInputStream(new ByteArrayInputStream(xlsx))) {
            java.util.zip.ZipEntry entry;
            while ((entry = zip.getNextEntry()) != null) {
                if (!entry.getName().contains("drawing") || !entry.getName().endsWith(".xml")) {
                    continue;
                }
                foundDrawing = true;
                String xml = new String(zip.readAllBytes(), java.nio.charset.StandardCharsets.UTF_8);
                assertTrue(xml.contains("srgbClr val=\"DC143C\""), xml);
                org.junit.jupiter.api.Assertions.assertFalse(
                        java.util.regex.Pattern.compile("srgbClr val=\"[0-9A-Fa-f]{8}\"").matcher(xml).find(),
                        xml);
            }
        }
        assertTrue(foundDrawing);
    }
}
