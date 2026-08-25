package com.excelreplace.excel;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Set;
import java.util.stream.Stream;

public final class ExcelFiles {
    private ExcelFiles() {
    }

    public static boolean isExcelFile(Path path) {
        if (path == null || !Files.isRegularFile(path)) {
            return false;
        }
        String name = path.getFileName().toString().toLowerCase(Locale.ROOT);
        if (name.startsWith("~$")) {
            return false;
        }
        return name.endsWith(".xlsx") || name.endsWith(".xlsm") || name.endsWith(".xls");
    }

    public static boolean isGeneratedOutput(Path path) {
        if (path == null) {
            return false;
        }
        String name = path.getFileName().toString();
        int dot = name.lastIndexOf('.');
        String stem = dot < 0 ? name : name.substring(0, dot);
        return stem.endsWith("_replaced");
    }

    public static Path defaultReplacedPath(Path input) {
        return withSuffix(input, "_replaced");
    }

    public static Path defaultDumpPath(Path excelFile) {
        return excelFile.resolveSibling(dumpFileName(excelFile.getFileName().toString()));
    }

    public static Path dumpPathBeside(Path nameSource, Path siblingExcel) {
        return siblingExcel.resolveSibling(dumpFileName(nameSource.getFileName().toString()));
    }

    /** 差分テキストは変更後 Excel と同じフォルダにだけ出力する。 */
    public static Path diffDumpPath(Path excelFile) {
        return excelFile.resolveSibling(diffDumpFileName(excelFile.getFileName().toString()));
    }

    public static void createParentDirectories(Path file) throws IOException {
        Path parent = file == null ? null : file.getParent();
        Files.createDirectories(parent == null ? Path.of(".") : parent);
    }

    static String dumpFileName(String excelName) {
        String stem = stem(excelName);
        String ext = extension(excelName);
        if (ext.startsWith(".")) {
            ext = ext.substring(1);
        }
        if (ext.isEmpty()) {
            return stem + ".txt";
        }
        return stem + "_" + ext.toLowerCase(Locale.ROOT) + ".txt";
    }

    static String diffDumpFileName(String excelName) {
        String stem = stem(excelName);
        String ext = extension(excelName);
        if (ext.startsWith(".")) {
            ext = ext.substring(1);
        }
        if (ext.isEmpty()) {
            return stem + "_diff.txt";
        }
        return stem + "_" + ext.toLowerCase(Locale.ROOT) + "_diff.txt";
    }

    static String stem(String name) {
        int dot = name.lastIndexOf('.');
        return dot < 0 ? name : name.substring(0, dot);
    }

    public static Path withSuffix(Path input, String suffix) {
        String name = input.getFileName().toString();
        return input.resolveSibling(stem(name) + suffix + extension(name));
    }

    static String extension(String name) {
        int dot = name.lastIndexOf('.');
        return dot < 0 ? "" : name.substring(dot);
    }

    public static List<Path> listExcelFiles(Path root, boolean recursive) throws IOException {
        return listExcelFiles(List.of(root), recursive);
    }

    public static List<Path> listExcelFiles(Collection<Path> roots, boolean recursive) throws IOException {
        Set<Path> files = new LinkedHashSet<>();
        for (Path root : roots) {
            Path normalized = root.toAbsolutePath().normalize();
            if (isExcelFile(normalized)) {
                files.add(normalized);
                continue;
            }
            if (!Files.isDirectory(normalized)) {
                continue;
            }
            int depth = recursive ? Integer.MAX_VALUE : 1;
            try (Stream<Path> stream = Files.walk(normalized, depth)) {
                stream.filter(ExcelFiles::isExcelFile)
                        .filter(path -> !isGeneratedOutput(path))
                        .map(path -> path.toAbsolutePath().normalize())
                        .forEach(files::add);
            }
        }
        List<Path> list = new ArrayList<>(files);
        list.sort(null);
        return list;
    }
}
