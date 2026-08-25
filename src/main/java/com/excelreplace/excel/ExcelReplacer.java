package com.excelreplace.excel;

import com.excelreplace.model.ProcessOptions;
import com.excelreplace.model.ProcessResult;
import com.excelreplace.model.ReplaceRule;
import org.apache.poi.hssf.usermodel.HSSFComment;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFShape;
import org.apache.poi.hssf.usermodel.HSSFShapeGroup;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Comment;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFShapeGroup;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;

import java.awt.Color;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.HashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.function.Consumer;

public final class ExcelReplacer {

    public ProcessResult processFile(
            Path input,
            Path output,
            List<ReplaceRule> rules,
            ProcessOptions options,
            Consumer<String> log) throws IOException {
        if (log == null) {
            log = ignored -> {
            };
        }
        log.accept("処理開始: " + input.getFileName());
        Path temp = output.resolveSibling(output.getFileName().toString() + ".replacing.tmp");
        ProcessResult result;
        try (InputStream in = Files.newInputStream(input);
             Workbook workbook = WorkbookFactory.create(in)) {
            result = processWorkbook(workbook, rules, options, log);
            ExcelFiles.createParentDirectories(output);
            try (OutputStream out = Files.newOutputStream(temp)) {
                workbook.write(out);
            }
        } catch (RuntimeException e) {
            Files.deleteIfExists(temp);
            throw e;
        }
        Files.move(temp, output, StandardCopyOption.REPLACE_EXISTING);
        result.addFile();
        log.accept("出力: " + output);
        log.accept(result.summary());
        return result;
    }

    public ProcessResult processWorkbook(
            Workbook workbook,
            List<ReplaceRule> rules,
            ProcessOptions options) {
        return processWorkbook(workbook, rules, options, ignored -> {
        });
    }

    public ProcessResult processWorkbook(
            Workbook workbook,
            List<ReplaceRule> rules,
            ProcessOptions options,
            Consumer<String> log) {
        List<RichTextEngine.CompiledRule> compiled = RichTextEngine.compile(rules, options);
        if (compiled.isEmpty()) {
            throw new IllegalArgumentException("有効な置換ルールがありません。");
        }
        FontSupport fonts = new FontSupport(workbook);
        ProcessResult result = new ProcessResult();

        if (options.isSheetNames()) {
            replaceSheetNames(workbook, compiled, options, result, log);
        }

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            Sheet sheet = workbook.getSheetAt(i);
            String sheetName = sheet.getSheetName();
            boolean globallyExcluded = options.isSheetExcluded(sheetName);
            List<RichTextEngine.CompiledRule> sheetRules = RichTextEngine.filterForSheet(
                    compiled, sheetName, globallyExcluded);
            if (sheetRules.isEmpty()) {
                if (globallyExcluded) {
                    log.accept("シート: " + sheetName + "（対象外のためスキップ）");
                } else {
                    log.accept("シート: " + sheetName + "（該当ルールなし）");
                }
                continue;
            }
            if (globallyExcluded) {
                log.accept("シート: " + sheetName + "（対象外だがルール指定により処理）");
            } else {
                log.accept("シート: " + sheetName);
            }
            if (options.isCells()) {
                replaceCells(sheet, sheetRules, fonts, result);
            }
            List<RichTextEngine.CompiledRule> nonCellRules = RichTextEngine.filterOutsideCells(sheetRules);
            if (options.isShapes() && !nonCellRules.isEmpty()) {
                replaceShapes(sheet, nonCellRules, fonts, result, log);
            }
            if (options.isComments()) {
                replaceComments(sheet, sheetRules, fonts, result);
            }
            if (options.isHeadersFooters() && !nonCellRules.isEmpty()) {
                replaceHeadersFooters(sheet, nonCellRules, result);
            }
        }
        return result;
    }

    private void replaceSheetNames(
            Workbook workbook,
            List<RichTextEngine.CompiledRule> compiled,
            ProcessOptions options,
            ProcessResult result,
            Consumer<String> log) {
        Set<String> used = new HashSet<>();
        int count = workbook.getNumberOfSheets();
        String[] names = new String[count];
        for (int i = 0; i < count; i++) {
            names[i] = workbook.getSheetName(i);
        }
        for (int i = 0; i < count; i++) {
            boolean globallyExcluded = options.isSheetExcluded(names[i]);
            List<RichTextEngine.CompiledRule> nameRules = RichTextEngine.filterOutsideCells(
                    RichTextEngine.filterForSheet(compiled, names[i], globallyExcluded));
            if (nameRules.isEmpty()) {
                used.add(names[i].toLowerCase(Locale.ROOT));
                continue;
            }
            RichTextEngine.ApplyResult applied = RichTextEngine.apply(
                    RichTextEngine.singleRun(names[i], null), nameRules);
            String next = applied.changed() ? applied.text() : names[i];
            next = uniqueSheetName(sanitizeSheetName(next), used);
            used.add(next.toLowerCase(Locale.ROOT));
            if (!next.equals(names[i])) {
                workbook.setSheetName(i, next);
                result.addSheetNameHits(Math.max(applied.replacementCount(), 1));
                log.accept("  シート名: " + names[i] + " -> " + next);
            }
        }
    }

    private void replaceCells(
            Sheet sheet,
            List<RichTextEngine.CompiledRule> compiled,
            FontSupport fonts,
            ProcessResult result) {
        for (Row row : sheet) {
            if (row == null) {
                continue;
            }
            for (Cell cell : row) {
                if (cell == null || cell.getCellType() != CellType.STRING) {
                    continue;
                }
                List<RichTextEngine.CompiledRule> cellRules = RichTextEngine.filterForCell(
                        compiled, cell.getRowIndex(), cell.getColumnIndex());
                if (cellRules.isEmpty()) {
                    continue;
                }
                Font baseFont = sheet.getWorkbook().getFontAt(cell.getCellStyle().getFontIndex());
                List<RichTextEngine.Run> runs = fonts.toRuns(cell.getRichStringCellValue(), baseFont);
                RichTextEngine.ApplyResult applied = RichTextEngine.apply(runs, cellRules);
                if (!applied.changed()) {
                    continue;
                }
                cell.setCellValue(fonts.toRichText(applied.runs()));
                result.addCellHits(applied.replacementCount());
            }
        }
    }

    private void replaceComments(
            Sheet sheet,
            List<RichTextEngine.CompiledRule> compiled,
            FontSupport fonts,
            ProcessResult result) {
        Map<CellAddress, ? extends Comment> comments = sheet.getCellComments();
        if (comments == null || comments.isEmpty()) {
            return;
        }
        List<CellAddress> addresses = new ArrayList<>(comments.keySet());
        addresses.sort(Comparator.comparingInt(CellAddress::getRow).thenComparingInt(CellAddress::getColumn));
        Font defaultFont = sheet.getWorkbook().getFontAt(0);
        for (CellAddress address : addresses) {
            Comment comment = comments.get(address);
            if (comment == null || comment.getString() == null) {
                continue;
            }
            List<RichTextEngine.CompiledRule> commentRules = RichTextEngine.filterForCell(
                    compiled, address.getRow(), address.getColumn());
            if (commentRules.isEmpty()) {
                continue;
            }
            List<RichTextEngine.Run> runs = fonts.toRuns(comment.getString(), defaultFont);
            RichTextEngine.ApplyResult applied = RichTextEngine.apply(runs, commentRules);
            if (!applied.changed()) {
                continue;
            }
            comment.setString(fonts.toRichText(applied.runs()));
            result.addCommentHits(applied.replacementCount());
        }
    }

    private void replaceHeadersFooters(
            Sheet sheet,
            List<RichTextEngine.CompiledRule> compiled,
            ProcessResult result) {
        Header header = sheet.getHeader();
        Footer footer = sheet.getFooter();
        result.addHeaderHits(replaceHeaderFooterPart(header::getLeft, header::setLeft, compiled));
        result.addHeaderHits(replaceHeaderFooterPart(header::getCenter, header::setCenter, compiled));
        result.addHeaderHits(replaceHeaderFooterPart(header::getRight, header::setRight, compiled));
        result.addHeaderHits(replaceHeaderFooterPart(footer::getLeft, footer::setLeft, compiled));
        result.addHeaderHits(replaceHeaderFooterPart(footer::getCenter, footer::setCenter, compiled));
        result.addHeaderHits(replaceHeaderFooterPart(footer::getRight, footer::setRight, compiled));
    }

    private int replaceHeaderFooterPart(
            java.util.function.Supplier<String> getter,
            java.util.function.Consumer<String> setter,
            List<RichTextEngine.CompiledRule> compiled) {
        String original = getter.get();
        if (original == null || original.isEmpty()) {
            return 0;
        }
        RichTextEngine.ApplyResult applied = RichTextEngine.apply(
                RichTextEngine.singleRun(original, null), compiled);
        if (!applied.changed()) {
            return 0;
        }
        setter.accept(applied.text());
        return applied.replacementCount();
    }

    private void replaceShapes(
            Sheet sheet,
            List<RichTextEngine.CompiledRule> compiled,
            FontSupport fonts,
            ProcessResult result,
            Consumer<String> log) {
        if (sheet instanceof XSSFSheet xssfSheet) {
            XSSFDrawing drawing = xssfSheet.getDrawingPatriarch();
            if (drawing == null) {
                return;
            }
            for (XSSFShape shape : drawing.getShapes()) {
                replaceXssfShape(shape, compiled, result, log);
            }
            return;
        }
        try {
            var patriarch = sheet.getDrawingPatriarch();
            if (patriarch instanceof HSSFPatriarch hssf) {
                Font defaultFont = sheet.getWorkbook().getFontAt(0);
                for (HSSFShape shape : hssf.getChildren()) {
                    replaceHssfShape(shape, compiled, fonts, defaultFont, result, log);
                }
            }
        } catch (Exception e) {
            log.accept("  図形の読み込みに失敗: " + e.getMessage());
        }
    }

    private void replaceXssfShape(
            XSSFShape shape,
            List<RichTextEngine.CompiledRule> compiled,
            ProcessResult result,
            Consumer<String> log) {
        if (shape instanceof XSSFShapeGroup group) {
            for (XSSFShape child : group) {
                replaceXssfShape(child, compiled, result, log);
            }
            return;
        }
        if (!(shape instanceof XSSFSimpleShape simple)) {
            return;
        }
        try {
            List<RichTextEngine.Run> runs = extractXssfShapeRuns(simple);
            if (runs.isEmpty()) {
                return;
            }
            RichTextEngine.ApplyResult applied = RichTextEngine.apply(runs, compiled);
            if (!applied.changed()) {
                return;
            }
            writeXssfShapeRuns(simple, applied.runs());
            result.addShapeHits(applied.replacementCount());
        } catch (Exception e) {
            log.accept("  図形の置換をスキップ: " + e.getMessage());
        }
    }

    private List<RichTextEngine.Run> extractXssfShapeRuns(XSSFSimpleShape shape) {
        List<RichTextEngine.Run> runs = new ArrayList<>();
        List<XSSFTextParagraph> paragraphs = shape.getTextParagraphs();
        if (paragraphs == null || paragraphs.isEmpty()) {
            String text = shape.getText();
            return RichTextEngine.singleRun(text, null);
        }
        for (int i = 0; i < paragraphs.size(); i++) {
            XSSFTextParagraph paragraph = paragraphs.get(i);
            List<XSSFTextRun> textRuns = paragraph.getTextRuns();
            if (textRuns != null) {
                for (XSSFTextRun textRun : textRuns) {
                    String text = textRun.getText();
                    if (text == null || text.isEmpty()) {
                        continue;
                    }
                    runs.add(new RichTextEngine.Run(text, ShapeRunStyle.from(textRun), false, null));
                }
            }
            if (i < paragraphs.size() - 1) {
                Object style = (textRuns == null || textRuns.isEmpty())
                        ? null
                        : ShapeRunStyle.from(textRuns.get(textRuns.size() - 1));
                runs.add(new RichTextEngine.Run("\n", style, false, null));
            }
        }
        if (runs.isEmpty()) {
            return RichTextEngine.singleRun(shape.getText(), null);
        }
        return runs;
    }

    private void writeXssfShapeRuns(XSSFSimpleShape shape, List<RichTextEngine.Run> runs) {
        shape.clearText();
        XSSFTextParagraph paragraph = shape.addNewTextParagraph();
        for (RichTextEngine.Run run : runs) {
            String text = run.text;
            int start = 0;
            while (start <= text.length()) {
                int newline = text.indexOf('\n', start);
                String piece = newline < 0 ? text.substring(start) : text.substring(start, newline);
                if (!piece.isEmpty()) {
                    XSSFTextRun textRun = paragraph.addNewTextRun();
                    textRun.setText(piece);
                    ShapeRunStyle.apply(textRun, run);
                }
                if (newline < 0) {
                    break;
                }
                paragraph = shape.addNewTextParagraph();
                start = newline + 1;
            }
        }
    }

    private void replaceHssfShape(
            HSSFShape shape,
            List<RichTextEngine.CompiledRule> compiled,
            FontSupport fonts,
            Font defaultFont,
            ProcessResult result,
            Consumer<String> log) {
        if (shape instanceof HSSFShapeGroup group) {
            for (HSSFShape child : group.getChildren()) {
                replaceHssfShape(child, compiled, fonts, defaultFont, result, log);
            }
            return;
        }
        if (shape instanceof HSSFComment) {
            return;
        }
        if (!(shape instanceof HSSFSimpleShape simple)) {
            return;
        }
        try {
            HSSFRichTextString current = simple.getString();
            if (current == null || current.getString() == null || current.getString().isEmpty()) {
                return;
            }
            List<RichTextEngine.Run> runs = fonts.toRuns(current, defaultFont);
            RichTextEngine.ApplyResult applied = RichTextEngine.apply(runs, compiled);
            if (!applied.changed()) {
                return;
            }
            simple.setString((HSSFRichTextString) fonts.toRichText(applied.runs()));
            result.addShapeHits(applied.replacementCount());
        } catch (Exception e) {
            log.accept("  図形の置換をスキップ: " + e.getMessage());
        }
    }

    static String sanitizeSheetName(String name) {
        if (name == null) {
            return "Sheet";
        }
        String cleaned = name.replaceAll("[\\\\/\\*\\?\\:\\[\\]]", "_").trim();
        if (cleaned.isEmpty()) {
            cleaned = "Sheet";
        }
        if (cleaned.length() > 31) {
            cleaned = cleaned.substring(0, 31);
        }
        return cleaned;
    }

    static String uniqueSheetName(String name, Set<String> usedLower) {
        String candidate = name;
        int n = 2;
        while (usedLower.contains(candidate.toLowerCase(Locale.ROOT))) {
            String suffix = "(" + n + ")";
            String base = name;
            if (base.length() + suffix.length() > 31) {
                base = base.substring(0, Math.max(1, 31 - suffix.length()));
            }
            candidate = base + suffix;
            n++;
        }
        return candidate;
    }

    static final class ShapeRunStyle {
        String fontFamily;
        Double fontSize;
        Boolean bold;
        Boolean italic;
        Boolean underline;
        Boolean strikeout;
        Color color;

        static ShapeRunStyle from(XSSFTextRun run) {
            ShapeRunStyle style = new ShapeRunStyle();
            try {
                var rPr = run.getXmlObject().isSetRPr() ? run.getXmlObject().getRPr() : null;
                if (rPr != null && rPr.isSetLatin() && rPr.getLatin().getTypeface() != null) {
                    style.fontFamily = rPr.getLatin().getTypeface();
                }
                if (rPr != null && rPr.isSetSz()) {
                    double size = run.getFontSize();
                    if (size > 0) {
                        style.fontSize = size;
                    }
                }
                if (rPr != null && rPr.isSetB() && run.isBold()) {
                    style.bold = true;
                }
                if (rPr != null && rPr.isSetI() && run.isItalic()) {
                    style.italic = true;
                }
                if (rPr != null && rPr.isSetU() && run.isUnderline()) {
                    style.underline = true;
                }
                if (rPr != null && rPr.isSetStrike() && run.isStrikethrough()) {
                    style.strikeout = true;
                }
                if (rPr != null && rPr.isSetSolidFill()) {
                    style.color = run.getFontColor();
                }
            } catch (Exception ignored) {
                // keep defaults when a run does not expose formatting
            }
            return style;
        }

        static void apply(XSSFTextRun textRun, RichTextEngine.Run run) {
            ShapeRunStyle style = run.fontKey instanceof ShapeRunStyle s ? s : null;
            if (style != null) {
                if (style.fontFamily != null && !style.fontFamily.isBlank()) {
                    textRun.setFont(style.fontFamily);
                }
                if (style.fontSize != null && style.fontSize > 0) {
                    textRun.setFontSize(style.fontSize);
                }
                if (Boolean.TRUE.equals(style.bold)) {
                    textRun.setBold(true);
                }
                if (Boolean.TRUE.equals(style.italic)) {
                    textRun.setItalic(true);
                }
                if (Boolean.TRUE.equals(style.underline)) {
                    textRun.setUnderline(true);
                }
                if (Boolean.TRUE.equals(style.strikeout)) {
                    textRun.setStrikethrough(true);
                }
            }
            Color color = run.replaced && run.replacementColor != null
                    ? run.replacementColor
                    : (style == null ? null : style.color);
            if (color != null) {
                textRun.setFontColor(color);
            }
        }
    }
}
