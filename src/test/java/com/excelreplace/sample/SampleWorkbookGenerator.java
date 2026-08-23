package com.excelreplace.sample;

import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Drawing;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.ShapeTypes;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * 手動確認用のサンプルブックを sample/ に出力する。
 */
public final class SampleWorkbookGenerator {
    public static void main(String[] args) throws Exception {
        Path dir = Path.of(args.length > 0 ? args[0] : "sample");
        Files.createDirectories(dir);
        Path xlsx = dir.resolve("設計書サンプル.xlsx");
        Path xls = dir.resolve("設計書サンプル.xls");
        try (XSSFWorkbook workbook = new XSSFWorkbook();
             OutputStream out = Files.newOutputStream(xlsx)) {
            fillXlsx(workbook);
            workbook.write(out);
        }
        try (HSSFWorkbook workbook = new HSSFWorkbook();
             OutputStream out = Files.newOutputStream(xls)) {
            fillXls(workbook);
            workbook.write(out);
        }
        System.out.println("作成: " + xlsx.toAbsolutePath());
        System.out.println("作成: " + xls.toAbsolutePath());
    }

    private static void fillXlsx(XSSFWorkbook workbook) {
        fillCommonSheets(workbook);
        XSSFSheet screen = workbook.getSheet("画面設計");
        XSSFDrawing drawing = screen.createDrawingPatriarch();
        XSSFClientAnchor anchor = drawing.createAnchor(0, 0, 0, 0, 1, 8, 4, 12);
        XSSFSimpleShape shape = drawing.createSimpleShape(anchor);
        shape.setShapeType(ShapeTypes.ROUND_RECT);
        shape.setLineStyleColor(0, 0, 0);
        shape.setFillColor(220, 230, 241);
        shape.setText("旧システム\nログイン");
    }

    private static void fillXls(HSSFWorkbook workbook) {
        fillCommonSheets(workbook);
        HSSFSheet screen = workbook.getSheet("画面設計");
        HSSFPatriarch patriarch = screen.createDrawingPatriarch();
        var anchor = patriarch.createAnchor(0, 0, 0, 0, 1, 8, 4, 12);
        HSSFSimpleShape shape = patriarch.createSimpleShape(anchor);
        shape.setShapeType(HSSFSimpleShape.OBJECT_TYPE_RECTANGLE);
        shape.setString(new HSSFRichTextString("旧システム\nログイン"));
    }

    private static void fillCommonSheets(Workbook workbook) {
        CellStyle title = titleStyle(workbook);
        CellStyle header = headerStyle(workbook);
        CellStyle body = bodyStyle(workbook);

        Sheet screen = workbook.createSheet("画面設計");
        screen.setColumnWidth(0, 18 * 256);
        screen.setColumnWidth(1, 28 * 256);
        screen.setColumnWidth(2, 22 * 256);
        screen.setColumnWidth(3, 22 * 256);
        screen.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
        row(screen, 0, title, "旧システム 画面設計書");
        row(screen, 1, header, "項目", "内容", "ID", "備考");
        row(screen, 2, body, "システム名", "旧システム", "ABC-001", "現行");
        row(screen, 3, body, "バージョン", "ver.1.0", "ABC-002", "2024年度");
        row(screen, 4, body, "画面名", "ログイン画面", "ABC-003", "旧システム 認証");
        row(screen, 5, body, "帳票名", "売上帳票", "ABC-004", "月次");
        addComment(screen, 2, 1, "旧システム の正式名称を確認すること");
        screen.getHeader().setCenter("旧システム 設計書");
        screen.getFooter().setLeft("旧システム");
        screen.getFooter().setRight("ver.1.0");

        Sheet list = workbook.createSheet("帳票一覧");
        list.setColumnWidth(0, 12 * 256);
        list.setColumnWidth(1, 24 * 256);
        list.setColumnWidth(2, 18 * 256);
        list.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
        row(list, 0, title, "旧システム 帳票一覧");
        row(list, 1, header, "No", "帳票名", "対象");
        row(list, 2, body, "1", "売上帳票", "旧システム");
        row(list, 3, body, "2", "在庫帳票", "旧システム");
        row(list, 4, body, "3", "ログイン履歴", "ABC-001");
        list.getHeader().setCenter("旧システム 帳票一覧");
    }

    private static void row(Sheet sheet, int rowIndex, CellStyle style, String... values) {
        Row row = sheet.createRow(rowIndex);
        row.setHeightInPoints(18);
        for (int i = 0; i < values.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(values[i]);
            cell.setCellStyle(style);
        }
    }

    private static CellStyle titleStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        font.setFontHeightInPoints((short) 14);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setAlignment(HorizontalAlignment.LEFT);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static CellStyle headerStyle(Workbook workbook) {
        Font font = workbook.createFont();
        font.setBold(true);
        CellStyle style = workbook.createCellStyle();
        style.setFont(font);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setAlignment(HorizontalAlignment.CENTER);
        applyBorder(style);
        return style;
    }

    private static CellStyle bodyStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        applyBorder(style);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    private static void applyBorder(CellStyle style) {
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
    }

    private static void addComment(Sheet sheet, int rowIndex, int colIndex, String text) {
        CreationHelper helper = sheet.getWorkbook().getCreationHelper();
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        ClientAnchor anchor = helper.createClientAnchor();
        anchor.setCol1(colIndex);
        anchor.setRow1(rowIndex);
        anchor.setCol2(colIndex + 3);
        anchor.setRow2(rowIndex + 4);
        Comment comment = drawing.createCellComment(anchor);
        comment.setAuthor("設計");
        comment.setString(helper.createRichTextString(text));
        sheet.getRow(rowIndex).getCell(colIndex).setCellComment(comment);
    }
}
