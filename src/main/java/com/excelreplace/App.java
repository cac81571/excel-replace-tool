package com.excelreplace;

import com.excelreplace.ui.MainFrame;
import com.formdev.flatlaf.FlatIntelliJLaf;

import javax.swing.SwingUtilities;
import javax.swing.UIManager;

public final class App {
    public static void main(String[] args) {
        System.setProperty("awt.useSystemAAFontSettings", "on");
        System.setProperty("swing.aatext", "true");
        FlatIntelliJLaf.setup();
        UIManager.put("Component.arrowType", "chevron");
        UIManager.put("Button.arc", 8);
        UIManager.put("TextComponent.arc", 8);
        SwingUtilities.invokeLater(() -> {
            MainFrame frame = new MainFrame();
            frame.setVisible(true);
        });
    }
}
