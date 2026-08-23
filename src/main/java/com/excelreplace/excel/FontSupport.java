package com.excelreplace.excel;

import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.awt.Color;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

final class FontSupport {
    private final Workbook workbook;
    private final Map<String, Font> cache = new HashMap<>();

    FontSupport(Workbook workbook) {
        this.workbook = workbook;
    }

    List<RichTextEngine.Run> toRuns(RichTextString rts, Font defaultFont) {
        if (rts == null) {
            return List.of();
        }
        String text = rts.getString();
        if (text == null || text.isEmpty()) {
            return List.of();
        }
        int n = rts.numFormattingRuns();
        if (n <= 0) {
            return RichTextEngine.singleRun(text, defaultFont);
        }
        if (rts instanceof XSSFRichTextString xssf) {
            return xssfRuns(xssf, text, defaultFont, n);
        }
        if (rts instanceof HSSFRichTextString hssf) {
            return hssfRuns(hssf, text, defaultFont, n);
        }
        return RichTextEngine.singleRun(text, defaultFont);
    }

    RichTextString toRichText(List<RichTextEngine.Run> runs) {
        String text = RichTextEngine.concat(runs);
        if (workbook instanceof XSSFWorkbook) {
            XSSFRichTextString rts = new XSSFRichTextString(text);
            apply(rts, runs);
            return rts;
        }
        HSSFRichTextString rts = new HSSFRichTextString(text);
        apply(rts, runs);
        return rts;
    }

    Font resolve(RichTextEngine.Run run, Font fallback) {
        Font base = run.fontKey instanceof Font font ? font : fallback;
        if (!run.replaced || run.replacementColor == null) {
            return base != null ? base : workbook.getFontAt(0);
        }
        String key = cacheKey(base, run.replacementColor);
        return cache.computeIfAbsent(key, ignored -> cloneWithColor(base, run.replacementColor));
    }

    private void apply(RichTextString rts, List<RichTextEngine.Run> runs) {
        int pos = 0;
        for (RichTextEngine.Run run : runs) {
            int end = pos + run.text.length();
            if (end > pos) {
                Font font = resolve(run, run.fontKey instanceof Font f ? f : workbook.getFontAt(0));
                rts.applyFont(pos, end, font);
            }
            pos = end;
        }
    }

    private List<RichTextEngine.Run> xssfRuns(XSSFRichTextString rts, String text, Font defaultFont, int n) {
        List<RichTextEngine.Run> runs = new ArrayList<>();
        int covered = 0;
        for (int i = 0; i < n; i++) {
            int start = rts.getIndexOfFormattingRun(i);
            if (start > covered) {
                runs.add(new RichTextEngine.Run(text.substring(covered, start), defaultFont, false, null));
            }
            int len = rts.getLengthOfFormattingRun(i);
            int end = Math.min(text.length(), start + Math.max(len, 0));
            Font font = rts.getFontOfFormattingRun(i);
            if (font == null) {
                font = defaultFont;
            }
            if (end > start) {
                runs.add(new RichTextEngine.Run(text.substring(start, end), font, false, null));
            }
            covered = Math.max(covered, end);
        }
        if (covered < text.length()) {
            runs.add(new RichTextEngine.Run(text.substring(covered), defaultFont, false, null));
        }
        return runs;
    }

    private List<RichTextEngine.Run> hssfRuns(HSSFRichTextString rts, String text, Font defaultFont, int n) {
        List<RichTextEngine.Run> runs = new ArrayList<>();
        int covered = 0;
        for (int i = 0; i < n; i++) {
            int start = rts.getIndexOfFormattingRun(i);
            if (start > covered) {
                runs.add(new RichTextEngine.Run(text.substring(covered, start), defaultFont, false, null));
            }
            int end = (i + 1 < n) ? rts.getIndexOfFormattingRun(i + 1) : text.length();
            Font font = defaultFont;
            short index = rts.getFontOfFormattingRun(i);
            try {
                font = workbook.getFontAt(index);
            } catch (Exception ignored) {
                font = defaultFont;
            }
            if (end > start) {
                runs.add(new RichTextEngine.Run(text.substring(start, Math.min(end, text.length())), font, false, null));
            }
            covered = Math.max(covered, end);
        }
        if (covered < text.length()) {
            runs.add(new RichTextEngine.Run(text.substring(covered), defaultFont, false, null));
        }
        return runs;
    }

    private Font cloneWithColor(Font base, Color color) {
        if (workbook instanceof XSSFWorkbook xssf) {
            XSSFFont font = xssf.createFont();
            copy(base, font);
            byte[] rgb = {(byte) color.getRed(), (byte) color.getGreen(), (byte) color.getBlue()};
            font.setColor(new XSSFColor(rgb, new DefaultIndexedColorMap()));
            return font;
        }
        HSSFWorkbook hssf = (HSSFWorkbook) workbook;
        HSSFFont font = hssf.createFont();
        copy(base, font);
        HSSFPalette palette = hssf.getCustomPalette();
        HSSFColor found = palette.findSimilarColor(color.getRed(), color.getGreen(), color.getBlue());
        if (found != null) {
            font.setColor(found.getIndex());
        }
        return font;
    }

    private static void copy(Font from, Font to) {
        if (from == null) {
            return;
        }
        to.setFontName(from.getFontName());
        to.setFontHeight(from.getFontHeight());
        to.setBold(from.getBold());
        to.setItalic(from.getItalic());
        to.setUnderline(from.getUnderline());
        to.setStrikeout(from.getStrikeout());
        to.setTypeOffset(from.getTypeOffset());
        to.setCharSet(from.getCharSet());
    }

    private static String cacheKey(Font base, Color color) {
        String name = base == null ? "" : base.getFontName();
        short height = base == null ? 0 : base.getFontHeight();
        boolean bold = base != null && base.getBold();
        boolean italic = base != null && base.getItalic();
        byte underline = base == null ? 0 : base.getUnderline();
        boolean strike = base != null && base.getStrikeout();
        return name + "|" + height + "|" + bold + "|" + italic + "|" + underline + "|" + strike
                + "|" + color.getRGB();
    }
}
