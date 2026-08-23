package com.excelreplace.ui;

import com.excelreplace.model.ReplaceRule;

import javax.swing.table.AbstractTableModel;
import java.util.ArrayList;
import java.util.List;

public final class RuleTableModel extends AbstractTableModel {
    private static final String[] COLUMNS = {"有効", "正規表現", "検索", "置換後"};
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
        return columnIndex <= 1 ? Boolean.class : String.class;
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
            case 2 -> rule.getPatternText();
            case 3 -> rule.getReplacement();
            default -> null;
        };
    }

    @Override
    public void setValueAt(Object aValue, int rowIndex, int columnIndex) {
        ReplaceRule rule = rules.get(rowIndex);
        switch (columnIndex) {
            case 0 -> rule.setEnabled(Boolean.TRUE.equals(aValue));
            case 1 -> rule.setRegex(Boolean.TRUE.equals(aValue));
            case 2 -> rule.setPatternText(aValue == null ? "" : aValue.toString());
            case 3 -> rule.setReplacement(aValue == null ? "" : aValue.toString());
            default -> {
            }
        }
        fireTableCellUpdated(rowIndex, columnIndex);
    }
}
