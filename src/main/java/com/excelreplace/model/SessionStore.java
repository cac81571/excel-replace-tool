package com.excelreplace.model;

import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;

/**
 * 前回起動時の入力値をユーザーホームに保存・復元する。
 */
public final class SessionStore {
    private SessionStore() {
    }

    public static Path file() {
        return Path.of(System.getProperty("user.home"), ".excel-replace-tool", "last-session.txt");
    }

    public static void save(AppSettings settings) throws IOException {
        Path path = file();
        Path parent = path.getParent();
        if (parent != null) {
            Files.createDirectories(parent);
        }
        Files.writeString(path, settings.format(), StandardCharsets.UTF_8);
    }

    public static AppSettings load() throws IOException {
        Path path = file();
        if (!Files.isRegularFile(path)) {
            return null;
        }
        return AppSettings.parse(Files.readString(path, StandardCharsets.UTF_8));
    }
}
