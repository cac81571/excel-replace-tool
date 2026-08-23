package com.excelreplace.model;

public final class ProcessResult {
    private int files;
    private int cellHits;
    private int shapeHits;
    private int commentHits;
    private int headerHits;
    private int sheetNameHits;

    public int getFiles() {
        return files;
    }

    public void addFile() {
        files++;
    }

    public int getCellHits() {
        return cellHits;
    }

    public void addCellHits(int n) {
        cellHits += n;
    }

    public int getShapeHits() {
        return shapeHits;
    }

    public void addShapeHits(int n) {
        shapeHits += n;
    }

    public int getCommentHits() {
        return commentHits;
    }

    public void addCommentHits(int n) {
        commentHits += n;
    }

    public int getHeaderHits() {
        return headerHits;
    }

    public void addHeaderHits(int n) {
        headerHits += n;
    }

    public int getSheetNameHits() {
        return sheetNameHits;
    }

    public void addSheetNameHits(int n) {
        sheetNameHits += n;
    }

    public int totalHits() {
        return cellHits + shapeHits + commentHits + headerHits + sheetNameHits;
    }

    public void merge(ProcessResult other) {
        files += other.files;
        cellHits += other.cellHits;
        shapeHits += other.shapeHits;
        commentHits += other.commentHits;
        headerHits += other.headerHits;
        sheetNameHits += other.sheetNameHits;
    }

    public String summary() {
        return String.format(
                "ファイル %d 件 / 置換 %d 件（セル %d, 図形 %d, コメント %d, ヘッダー %d, シート名 %d）",
                files, totalHits(), cellHits, shapeHits, commentHits, headerHits, sheetNameHits);
    }
}
