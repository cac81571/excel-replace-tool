package com.excelreplace.ui;

import com.excelreplace.model.ProcessOptions;
import com.excelreplace.model.ReplaceRule;

import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.List;

public final class RuleTableModel extends AbstractTableModel {
    private static final String[] COLUMNS = {
            "有効", "正規表現", "大/小無視", "検索", "置換後", "対象シート", "セル範囲"
    };
    private final List<ReplaceRule> rules = new ArrayList<>();

    public RuleTableModel() {
        addRule();
    }

    public List<ReplaceRule> getRules() {
        return rules;
    }

    public void addRule() {
        rules.add(new ReplaceRule());
        int row = rules.size() - 1;
        fireTableRowsInserted(row, row);
    }

    public void removeRows(int[] rows) {
        List<Integer> sorted = new ArrayList<>();
        for (int row : rows) {
            sorted.add(row);
        }
        sorted.sort((a, b) -> Integer.compare(b, a));
        for (int row : sorted) {
            if (row >= 0 && row < rules.size()) {
                rules.remove(row);
                fireTableRowsDeleted(row, row);
            }
        }
        if (rules.isEmpty()) {
            addRule();
        }
    }

    /**
     * 選択行を 1 つ上へ。移動後の選択行インデックスを返す（移動不可なら空）。
     */
    public int[] moveRowsUp(int[] rows) {
        List<Integer> selected = sortedUnique(rows);
        if (selected.isEmpty() || selected.get(0) <= 0) {
            return new int[0];
        }
        for (int row : selected) {
            ReplaceRule rule = rules.remove(row);
            rules.add(row - 1, rule);
        }
        fireTableDataChanged();
        return selected.stream().mapToInt(row -> row - 1).toArray();
    }

    /**
     * 選択行を 1 つ下へ。移動後の選択行インデックスを返す（移動不可なら空）。
     */
    public int[] moveRowsDown(int[] rows) {
        List<Integer> selected = sortedUnique(rows);
        if (selected.isEmpty() || selected.get(selected.size() - 1) >= rules.size() - 1) {
            return new int[0];
        }
        for (int i = selected.size() - 1; i >= 0; i--) {
            int row = selected.get(i);
            ReplaceRule rule = rules.remove(row);
            rules.add(row + 1, rule);
        }
        fireTableDataChanged();
        return selected.stream().mapToInt(row -> row + 1).toArray();
    }

    private static List<Integer> sortedUnique(int[] rows) {
        List<Integer> selected = new ArrayList<>();
        if (rows == null) {
            return selected;
        }
        for (int row : rows) {
            if (row >= 0 && !selected.contains(row)) {
                selected.add(row);
            }
        }
        selected.sort(Integer::compareTo);
        return selected;
    }

    public List<ReplaceRule> snapshot() {
        return snapshot(new int[0]);
    }

    public void replaceAll(List<ReplaceRule> next) {
        rules.clear();
        if (next == null || next.isEmpty()) {
            rules.add(new ReplaceRule());
        } else {
            for (ReplaceRule rule : next) {
                rules.add(rule.copy());
            }
        }
        fireTableDataChanged();
    }

    public List<ReplaceRule> snapshot(int[] rows) {
        List<ReplaceRule> selected = new ArrayList<>();
        if (rows == null || rows.length == 0) {
            for (ReplaceRule rule : rules) {
                if (!rule.getPatternText().isBlank() || !rule.getReplacement().isBlank()) {
                    selected.add(rule.copy());
                }
            }
            return selected;
        }
        for (int row : rows) {
            if (row >= 0 && row < rules.size()) {
                selected.add(rules.get(row).copy());
            }
        }
        return selected;
    }

    @Override
    public int getRowCount() {
        return rules.size();
    }

    @Override
    public int getColumnCount() {
        return COLUMNS.length;
    }

    @Override
    public String getColumnName(int column) {
        return COLUMNS[column];
    }

    @Override
    public Class<?> getColumnClass(int columnIndex) {
        return columnIndex <= 2 ? Boolean.class : String.class;
    }

    @Override
    public boolean isCellEditable(int rowIndex, int columnIndex) {
        return true;
    }

    @Override
    public Object getValueAt(int rowIndex, int columnIndex) {
        ReplaceRule rule = rules.get(rowIndex);
        return switch (columnIndex) {
            case 0 -> rule.isEnabled();
            case 1 -> rule.isRegex();
            case 2 -> rule.isIgnoreCase();
            case 3 -> rule.getPatternText();
            case 4 -> rule.getReplacement();
            case 5 -> ProcessOptions.formatSheetList(rule.getTargetSheets());
            case 6 -> ReplaceRule.formatRangeList(rule.getCellRanges());
            default -> null;
        };
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        ReplaceRule rule = rules.get(rowIndex);
        String text = aValue == null ? "" : aValue.toString();
        switch (columnIndex) {
            case 0 -> rule.setEnabled(Boolean.TRUE.equals(aValue));
            case 1 -> rule.setRegex(Boolean.TRUE.equals(aValue));
            case 2 -> rule.setIgnoreCase(Boolean.TRUE.equals(aValue));
            case 3 -> rule.setPatternText(text);
            case 4 -> rule.setReplacement(text);
            case 5 -> rule.setTargetSheets(ProcessOptions.parseSheetList(text));
            case 6 -> rule.setCellRanges(ReplaceRule.parseRangeList(text));
            default -> {
            }
        }
        fireTableCellUpdated(rowIndex, columnIndex);
    }
}
