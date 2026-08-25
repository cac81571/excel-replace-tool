package com.excelreplace.excel;

import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFAnchor;
import org.apache.poi.xssf.usermodel.XSSFClientAnchor;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;

import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.Map;

/**
 * Excel の内容を 1 行 1 レコードのテキストに落とし、WinMerge 等での DIFF を安定させる。
 * 内容行は先頭列にシート名を置き、シートをまたいでも位置が分かるようにする。
 */
public final class ExcelTextDumper {
    private final DataFormatter formatter = new DataFormatter();

    public String dumpFile(Path file) throws IOException {
        try (InputStream in = Files.newInputStream(file);
             Workbook workbook = WorkbookFactory.create(in)) {
            return dump(workbook, file.getFileName().toString());
        }
    }

    public String dump(Workbook workbook) {
        return dump(workbook, "");
    }

    public String dump(Workbook workbook, String fileName) {
        StringBuilder out = new StringBuilder();
        out.append("FILE\t").append(esc(fileName)).append('\n');
        out.append("SHEETS\t").append(workbook.getNumberOfSheets()).append('\n');
        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            dumpSheet(out, i, sheet);
        }
        return out.toString();
    }

    private void dumpSheet(StringBuilder out, int index, Sheet sheet) {
        var workbook = sheet.getWorkbook();
        boolean hidden = workbook.isSheetHidden(index) || workbook.isSheetVeryHidden(index);
        String sheetName = sheet.getSheetName();
        out.append("SHEET\t").append(index)
                .append('\t').append(esc(sheetName))
                .append('\t').append(hidden ? "hidden" : "visible")
                .append('\n');

        for (int m = 0; m < sheet.getNumMergedRegions(); m++) {
            CellRangeAddress range = sheet.getMergedRegion(m);
            out.append(esc(sheetName)).append("\tMERGE\t").append(range.formatAsString()).append('\n');
        }

        dumpHeaderFooter(out, sheetName, "HEADER", sheet.getHeader());
        dumpHeaderFooter(out, sheetName, "FOOTER", sheet.getFooter());

        for (Row row : sheet) {
            if (row == null) {
                continue;
            }
            for (Cell cell : row) {
                dumpCell(out, sheetName, cell);
            }
        }

        Map<CellAddress, ? extends Comment> comments = sheet.getCellComments();
        if (comments != null && !comments.isEmpty()) {
            List<CellAddress> addresses = new ArrayList<>(comments.keySet());
            addresses.sort(Comparator.comparingInt(CellAddress::getRow).thenComparingInt(CellAddress::getColumn));
            for (CellAddress address : addresses) {
                Comment comment = comments.get(address);
                if (comment == null || comment.getString() == null) {
                    continue;
                }
                out.append(esc(sheetName)).append("\tCOMMENT\t").append(address.formatAsString())
                        .append('\t').append(esc(comment.getAuthor()))
                        .append('\t').append(esc(comment.getString().getString()))
                        .append('\n');
            }
        }

        dumpShapes(out, sheetName, sheet);
    }

    private void dumpCell(StringBuilder out, String sheetName, Cell cell) {
        if (cell == null) {
            return;
        }
        CellType type = cell.getCellType();
        if (type == CellType.BLANK) {
            return;
        }
        String address = new CellAddress(cell).formatAsString();
        if (type == CellType.FORMULA) {
            CellType cached = cell.getCachedFormulaResultType();
            out.append(esc(sheetName)).append("\tFORMULA\t").append(address)
                    .append('\t').append(cached.name())
                    .append('\t').append(esc(cell.getCellFormula()))
                    .append('\t').append(esc(formatter.formatCellValue(cell)))
                    .append('\n');
            return;
        }
        out.append(esc(sheetName)).append("\tCELL\t").append(address)
                .append('\t').append(type.name())
                .append('\t').append(esc(plainValue(cell)))
                .append('\n');
    }

    private String plainValue(Cell cell) {
        return switch (cell.getCellType()) {
            case STRING -> cell.getRichStringCellValue().getString();
            case BOOLEAN -> Boolean.toString(cell.getBooleanCellValue());
            case ERROR -> "#ERROR";
            default -> formatter.formatCellValue(cell);
        };
    }

    private void dumpHeaderFooter(StringBuilder out, String sheetName, String kind, Header header) {
        dumpPart(out, sheetName, kind, "LEFT", header.getLeft());
        dumpPart(out, sheetName, kind, "CENTER", header.getCenter());
        dumpPart(out, sheetName, kind, "RIGHT", header.getRight());
    }

    private void dumpHeaderFooter(StringBuilder out, String sheetName, String kind, Footer footer) {
        dumpPart(out, sheetName, kind, "LEFT", footer.getLeft());
        dumpPart(out, sheetName, kind, "CENTER", footer.getCenter());
        dumpPart(out, sheetName, kind, "RIGHT", footer.getRight());
    }

    private void dumpPart(StringBuilder out, String sheetName, String kind, String pos, String value) {
        if (value == null || value.isEmpty()) {
            return;
        }
        out.append(esc(sheetName)).append('\t').append(kind).append('\t')
                .append(pos).append('\t').append(esc(value)).append('\n');
    }

    private void dumpShapes(StringBuilder out, String sheetName, Sheet sheet) {
        if (sheet instanceof XSSFSheet xssfSheet) {
            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            if (drawing == null) {
                return;
            }
            int[] index = {0};
            for (XSSFShape shape : drawing.getShapes()) {
                dumpXssfShape(out, sheetName, shape, index);
            }
            return;
        }
        try {
            var patriarch = sheet.getDrawingPatriarch();
            if (patriarch instanceof HSSFPatriarch hssf) {
                int[] index = {0};
                for (HSSFShape shape : hssf.getChildren()) {
                    dumpHssfShape(out, sheetName, shape, index);
                }
            }
        } catch (Exception ignored) {
            // 図形が読めないファイルはセルだけ残す
        }
    }

    private void dumpXssfShape(StringBuilder out, String sheetName, XSSFShape shape, int[] index) {
        if (shape instanceof XSSFShapeGroup group) {
            for (XSSFShape child : group) {
                dumpXssfShape(out, sheetName, child, index);
            }
            return;
        }
        if (!(shape instanceof XSSFSimpleShape simple)) {
            return;
        }
        String text;
        try {
            text = simple.getText();
        } catch (Exception e) {
            return;
        }
        if (text == null || text.isEmpty()) {
            return;
        }
        out.append(esc(sheetName)).append("\tSHAPE\t").append(index[0]++)
                .append('\t').append(anchorOf(shape))
                .append('\t').append(esc(text))
                .append('\n');
    }

    private void dumpHssfShape(StringBuilder out, String sheetName, HSSFShape shape, int[] index) {
        if (shape instanceof HSSFShapeGroup group) {
            for (HSSFShape child : group.getChildren()) {
                dumpHssfShape(out, sheetName, child, index);
            }
            return;
        }
        if (shape instanceof HSSFComment) {
            return;
        }
        if (!(shape instanceof HSSFSimpleShape simple)) {
            return;
        }
        HSSFRichTextString rts = simple.getString();
        if (rts == null || rts.getString() == null || rts.getString().isEmpty()) {
            return;
        }
        out.append(esc(sheetName)).append("\tSHAPE\t").append(index[0]++)
                .append('\t').append(anchorOf(shape))
                .append('\t').append(esc(rts.getString()))
                .append('\n');
    }

    private static String anchorOf(XSSFShape shape) {
        XSSFAnchor anchor = shape.getAnchor();
        if (anchor instanceof XSSFClientAnchor client) {
            return client.getCol1() + "," + client.getRow1() + ":" + client.getCol2() + "," + client.getRow2();
        }
        return "-";
    }

    private static String anchorOf(HSSFShape shape) {
        if (shape.getAnchor() instanceof HSSFClientAnchor client) {
            return client.getCol1() + "," + client.getRow1() + ":" + client.getCol2() + "," + client.getRow2();
        }
        return "-";
    }

    static String esc(String value) {
        if (value == null) {
            return "";
        }
        return value
                .replace("\\", "\\\\")
                .replace("\r", "\\r")
                .replace("\n", "\\n")
                .replace("\t", "\\t");
    }
}
