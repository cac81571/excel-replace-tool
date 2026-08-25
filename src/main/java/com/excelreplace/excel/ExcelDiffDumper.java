package com.excelreplace.excel;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

/**
 * 変換前後のテキストダンプを比較し、変更された箇所だけを出力する。
 * 各行は位置情報のあとに変更前・変更後の値を並べる。
 */
public final class ExcelDiffDumper {
    private ExcelDiffDumper() {
    }

    public static String diff(String beforeDump, String afterDump, String fileName) {
        Parsed before = parse(beforeDump);
        Parsed after = parse(afterDump);
        Set<String> keys = new LinkedHashSet<>();
        keys.addAll(after.records.keySet());
        keys.addAll(before.records.keySet());

        List<String> lines = new ArrayList<>();
        for (String key : keys) {
            Record left = before.records.get(key);
            Record right = after.records.get(key);
            String beforeValue = left == null ? "" : left.value;
            String afterValue = right == null ? "" : right.value;
            if (beforeValue.equals(afterValue)) {
                continue;
            }
            String identity = right != null ? right.identity : left.identity;
            lines.add(identity + '\t' + beforeValue + '\t' + afterValue);
        }

        StringBuilder out = new StringBuilder();
        out.append("FILE\t").append(ExcelTextDumper.esc(fileName == null ? "" : fileName)).append('\n');
        out.append("CHANGED\t").append(lines.size()).append('\n');
        for (String line : lines) {
            out.append(line).append('\n');
        }
        return out.toString();
    }

    private static Parsed parse(String dump) {
        Parsed parsed = new Parsed();
        if (dump == null || dump.isBlank()) {
            return parsed;
        }
        Map<String, Integer> sheetIndexByName = new LinkedHashMap<>();
        String[] rawLines = dump.replace("\r\n", "\n").replace('\r', '\n').split("\n", -1);
        for (String line : rawLines) {
            if (line.isEmpty()) {
                continue;
            }
            String[] cols = line.split("\t", -1);
            if (cols.length == 0) {
                continue;
            }
            if ("FILE".equals(cols[0]) || "SHEETS".equals(cols[0])) {
                continue;
            }
            if ("SHEET".equals(cols[0]) && cols.length >= 3) {
                int index = parseInt(cols[1], -1);
                String name = cols[2];
                if (index >= 0) {
                    sheetIndexByName.put(name, index);
                    String visibility = cols.length >= 4 ? cols[3] : "";
                    String key = "SHEET\t" + index;
                    parsed.records.put(key, new Record(key, key, name + '\t' + visibility));
                }
                continue;
            }
            if (cols.length < 2) {
                continue;
            }
            String sheetName = cols[0];
            Integer sheetIndex = sheetIndexByName.get(sheetName);
            int index = sheetIndex == null ? -1 : sheetIndex;
            Record record = toRecord(sheetName, index, cols[1], cols);
            if (record != null) {
                parsed.records.put(record.key, record);
            }
        }
        return parsed;
    }

    private static Record toRecord(String sheetName, int sheetIndex, String kind, String[] cols) {
        return switch (kind) {
            case "CELL" -> {
                if (cols.length < 5) {
                    yield null;
                }
                String locator = cols[2] + '\t' + cols[3];
                String identity = sheetName + "\tCELL\t" + locator;
                yield new Record(key(sheetIndex, "CELL", locator), identity, cols[4]);
            }
            case "FORMULA" -> {
                if (cols.length < 4) {
                    yield null;
                }
                String locator = cols[2];
                String identity = sheetName + "\tFORMULA\t" + locator;
                yield new Record(key(sheetIndex, "FORMULA", locator), identity, join(cols, 3));
            }
            case "COMMENT" -> {
                if (cols.length < 4) {
                    yield null;
                }
                String locator = cols[2];
                String identity = sheetName + "\tCOMMENT\t" + locator;
                yield new Record(key(sheetIndex, "COMMENT", locator), identity, join(cols, 3));
            }
            case "SHAPE" -> {
                if (cols.length < 5) {
                    yield null;
                }
                String locator = cols[2] + '\t' + cols[3];
                String identity = sheetName + "\tSHAPE\t" + locator;
                yield new Record(key(sheetIndex, "SHAPE", locator), identity, cols[4]);
            }
            case "HEADER", "FOOTER" -> {
                if (cols.length < 4) {
                    yield null;
                }
                String locator = cols[2];
                String identity = sheetName + '\t' + kind + '\t' + locator;
                yield new Record(key(sheetIndex, kind, locator), identity, cols[3]);
            }
            case "MERGE" -> {
                if (cols.length < 3) {
                    yield null;
                }
                String locator = cols[2];
                String identity = sheetName + "\tMERGE\t" + locator;
                yield new Record(key(sheetIndex, "MERGE", locator), identity, "");
            }
            default -> null;
        };
    }

    private static String key(int sheetIndex, String kind, String locator) {
        return sheetIndex + "\t" + kind + '\t' + locator;
    }

    private static String join(String[] cols, int from) {
        StringBuilder out = new StringBuilder();
        for (int i = from; i < cols.length; i++) {
            if (i > from) {
                out.append('\t');
            }
            out.append(cols[i]);
        }
        return out.toString();
    }

    private static int parseInt(String text, int fallback) {
        try {
            return Integer.parseInt(text);
        } catch (NumberFormatException e) {
            return fallback;
        }
    }

    private static final class Parsed {
        private final Map<String, Record> records = new LinkedHashMap<>();
    }

    private static final class Record {
        private final String key;
        private final String identity;
        private final String value;

        private Record(String key, String identity, String value) {
            this.key = key;
            this.identity = identity;
            this.value = value;
        }
    }
}
