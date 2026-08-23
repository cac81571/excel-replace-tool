package com.excelreplace.ui;

import com.excelreplace.excel.ExcelFiles;
import com.excelreplace.excel.ExcelReplacer;
import com.excelreplace.excel.ExcelTextDumper;
import com.excelreplace.excel.RichTextEngine;
import com.excelreplace.model.AppSettings;
import com.excelreplace.model.ProcessOptions;
import com.excelreplace.model.ProcessResult;
import com.excelreplace.model.ReplaceRule;
import com.excelreplace.model.SessionStore;
import org.apache.poi.EncryptedDocumentException;

import javax.swing.BorderFactory;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JColorChooser;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.ListSelectionModel;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.Insets;
import java.awt.datatransfer.DataFlavor;
import java.awt.dnd.DnDConstants;
import java.awt.dnd.DropTarget;
import java.awt.dnd.DropTargetAdapter;
import java.awt.dnd.DropTargetDropEvent;
import java.awt.event.ActionListener;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;
import java.util.prefs.Preferences;

public final class MainFrame extends JFrame {
    private static final DateTimeFormatter TIME = DateTimeFormatter.ofPattern("HH:mm:ss");

    private final JTextField inputField = new JTextField();
    private final JTextField outputField = new JTextField();
    private final JCheckBox recursiveCheck = new JCheckBox("サブフォルダも含める");
    private final JCheckBox dumpTextCheck = new JCheckBox("テキスト出力（変換前・変換後）", true);
    private final JCheckBox cellsCheck = new JCheckBox("セル", true);
    private final JCheckBox shapesCheck = new JCheckBox("図形（オートシェイプ）", true);
    private final JCheckBox commentsCheck = new JCheckBox("コメント", true);
    private final JCheckBox headersCheck = new JCheckBox("ヘッダー/フッター", true);
    private final JCheckBox sheetNamesCheck = new JCheckBox("シート名");
    private final JCheckBox caseCheck = new JCheckBox("大文字小文字を無視");
    private final JCheckBox multilineCheck = new JCheckBox("複数行モード (^ $ を行単位)", true);
    private final JCheckBox recolorCheck = new JCheckBox("置換後の文字色を変更", true);
    private final JButton colorButton = new JButton();
    private Color replacementColor = new Color(220, 20, 60);
    private final RuleTableModel ruleModel = new RuleTableModel();
    private final JTable ruleTable = new JTable(ruleModel);
    private final JTextArea logArea = new JTextArea();
    private final JButton replaceButton = new JButton("置換を実行");
    private final Preferences prefs = Preferences.userNodeForPackage(MainFrame.class);
    private final ExcelReplacer replacer = new ExcelReplacer();
    private final ExcelTextDumper dumper = new ExcelTextDumper();

    public MainFrame() {
        super("Excel 設計書 置換ツール");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setMinimumSize(new Dimension(960, 680));
        setLocationByPlatform(true);

        JPanel root = new JPanel(new BorderLayout(8, 8));
        root.setBorder(BorderFactory.createEmptyBorder(12, 12, 12, 12));
        root.add(buildNorth(), BorderLayout.NORTH);
        root.add(buildCenter(), BorderLayout.CENTER);
        root.add(buildSouth(), BorderLayout.SOUTH);
        setContentPane(root);

        enableInputDrop();
        enableOutputDrop();
        configureRuleTable();
        restoreLastSession();
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                saveLastSessionQuietly();
            }
        });
    }

    private JPanel buildNorth() {
        JPanel panel = new JPanel(new GridBagLayout());
        GridBagConstraints c = new GridBagConstraints();
        c.insets = new Insets(4, 4, 4, 4);
        c.fill = GridBagConstraints.HORIZONTAL;
        c.gridy = 0;
        c.gridx = 0;
        c.weightx = 0;
        panel.add(new JLabel("入力（複数ファイル・フォルダ可）"), c);
        c.gridx = 1;
        c.weightx = 1;
        panel.add(inputField, c);
        c.gridx = 2;
        c.weightx = 0;
        panel.add(button("ファイル...", e -> chooseInputFiles()), c);
        c.gridx = 3;
        panel.add(button("フォルダ...", e -> chooseInputFolder()), c);

        c.gridy = 1;
        c.gridx = 0;
        panel.add(new JLabel("出力フォルダ（空なら *_replaced）"), c);
        c.gridx = 1;
        c.weightx = 1;
        panel.add(outputField, c);
        c.gridx = 2;
        c.weightx = 0;
        panel.add(button("フォルダ...", e -> chooseOutputFolder()), c);
        c.gridx = 3;
        panel.add(recursiveCheck, c);

        c.gridy = 2;
        c.gridx = 0;
        c.gridwidth = 4;
        JPanel targets = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 0));
        targets.setBorder(BorderFactory.createTitledBorder("置換対象"));
        targets.add(cellsCheck);
        targets.add(shapesCheck);
        targets.add(commentsCheck);
        targets.add(headersCheck);
        targets.add(sheetNamesCheck);
        targets.add(caseCheck);
        targets.add(multilineCheck);
        panel.add(targets, c);
        return panel;
    }

    private JPanel buildCenter() {
        JPanel panel = new JPanel(new BorderLayout(8, 8));

        JPanel rules = new JPanel(new BorderLayout(4, 4));
        rules.setBorder(BorderFactory.createTitledBorder(
                "置換ルール（上から順に適用。正規表現OFFなら文字列そのまま）"));
        ruleTable.setRowHeight(26);
        ruleTable.setSelectionMode(ListSelectionModel.MULTIPLE_INTERVAL_SELECTION);
        ruleTable.putClientProperty("terminateEditOnFocusLost", Boolean.TRUE);
        rules.add(new JScrollPane(ruleTable), BorderLayout.CENTER);

        JPanel ruleButtons = new JPanel(new BorderLayout());
        JPanel ruleEdit = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 4));
        ruleEdit.add(button("ルール追加", e -> ruleModel.addRule()));
        ruleEdit.add(button("選択行を削除", e -> ruleModel.removeRows(ruleTable.getSelectedRows())));
        ruleEdit.add(button("設定エクスポート", e -> exportSettings()));
        ruleEdit.add(button("設定インポート", e -> importSettings()));
        JPanel colorPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 4));
        colorPanel.add(recolorCheck);
        colorPanel.add(new JLabel("文字色"));
        colorButton.setPreferredSize(new Dimension(120, 28));
        colorButton.addActionListener(e -> chooseReplacementColor());
        recolorCheck.addActionListener(e -> {
            colorButton.setEnabled(recolorCheck.isSelected());
            updateColorButton();
        });
        updateColorButton();
        colorPanel.add(colorButton);
        ruleButtons.add(ruleEdit, BorderLayout.WEST);
        ruleButtons.add(colorPanel, BorderLayout.EAST);
        rules.add(ruleButtons, BorderLayout.SOUTH);

        JPanel log = new JPanel(new BorderLayout());
        log.setBorder(BorderFactory.createTitledBorder("ログ"));
        logArea.setEditable(false);
        logArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        logArea.setLineWrap(true);
        log.add(new JScrollPane(logArea), BorderLayout.CENTER);

        panel.add(rules, BorderLayout.CENTER);
        panel.add(log, BorderLayout.SOUTH);
        log.setPreferredSize(new Dimension(100, 180));
        return panel;
    }

    private JPanel buildSouth() {
        JPanel panel = new JPanel(new BorderLayout());
        JPanel left = new JPanel(new FlowLayout(FlowLayout.LEFT, 8, 4));
        left.add(dumpTextCheck);
        JPanel right = new JPanel(new FlowLayout(FlowLayout.RIGHT, 8, 4));
        right.add(replaceButton);
        replaceButton.addActionListener(e -> runReplace());
        panel.add(left, BorderLayout.WEST);
        panel.add(right, BorderLayout.EAST);
        return panel;
    }

    private void configureRuleTable() {
        ruleTable.getColumnModel().getColumn(0).setMaxWidth(50);
        ruleTable.getColumnModel().getColumn(0).setPreferredWidth(50);
        ruleTable.getColumnModel().getColumn(1).setMaxWidth(80);
        ruleTable.getColumnModel().getColumn(1).setPreferredWidth(80);
    }

    private AppSettings currentSettings() {
        stopEditing();
        AppSettings settings = new AppSettings();
        settings.setRecursive(recursiveCheck.isSelected());
        settings.setDumpText(dumpTextCheck.isSelected());
        settings.setInputPath(inputField.getText().trim());
        settings.setOutputPath(outputField.getText().trim());
        ProcessOptions options = currentOptions();
        ProcessOptions dest = settings.getOptions();
        dest.setCells(options.isCells());
        dest.setShapes(options.isShapes());
        dest.setComments(options.isComments());
        dest.setHeadersFooters(options.isHeadersFooters());
        dest.setSheetNames(options.isSheetNames());
        dest.setCaseInsensitive(options.isCaseInsensitive());
        dest.setMultiline(options.isMultiline());
        dest.setRecolor(options.isRecolor());
        dest.setReplacementColor(options.getReplacementColor());
        settings.getRules().addAll(ruleModel.snapshot());
        return settings;
    }

    private void applySettings(AppSettings settings) {
        ProcessOptions options = settings.getOptions();
        recursiveCheck.setSelected(settings.isRecursive());
        dumpTextCheck.setSelected(settings.isDumpText());
        inputField.setText(settings.getInputPath());
        outputField.setText(settings.getOutputPath());
        cellsCheck.setSelected(options.isCells());
        shapesCheck.setSelected(options.isShapes());
        commentsCheck.setSelected(options.isComments());
        headersCheck.setSelected(options.isHeadersFooters());
        sheetNamesCheck.setSelected(options.isSheetNames());
        caseCheck.setSelected(options.isCaseInsensitive());
        multilineCheck.setSelected(options.isMultiline());
        recolorCheck.setSelected(options.isRecolor());
        replacementColor = options.getReplacementColor();
        updateColorButton();
        ruleModel.replaceAll(settings.getRules());
    }

    private void restoreLastSession() {
        try {
            AppSettings settings = SessionStore.load();
            if (settings == null) {
                return;
            }
            applySettings(settings);
            appendLog("前回の入力を復元しました。");
        } catch (Exception e) {
            appendLog("前回入力の復元に失敗しました: " + e.getMessage());
        }
    }

    private void saveLastSessionQuietly() {
        try {
            SessionStore.save(currentSettings());
        } catch (Exception ignored) {
            // 終了時などに失敗しても処理は続ける
        }
    }

    private void exportSettings() {
        AppSettings settings = currentSettings();
        if (settings.getRules().isEmpty()) {
            int answer = JOptionPane.showConfirmDialog(
                    this,
                    "有効な置換ルールがありません。このまま設定を書き出しますか？",
                    "設定エクスポート",
                    JOptionPane.YES_NO_OPTION);
            if (answer != JOptionPane.YES_OPTION) {
                return;
            }
        }
        JFileChooser chooser = settingsChooser();
        chooser.setDialogTitle("設定のエクスポート");
        chooser.setSelectedFile(new File("excel-replace-settings.txt"));
        if (chooser.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File file = withSettingsExtension(chooser.getSelectedFile());
        try {
            Files.writeString(file.toPath(), settings.format(), StandardCharsets.UTF_8);
            remember(file);
            appendLog("設定を書き出しました: " + file.getAbsolutePath());
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "書き出しに失敗しました: " + e.getMessage(), "設定エクスポート", JOptionPane.ERROR_MESSAGE);
        }
    }

    private void importSettings() {
        JFileChooser chooser = settingsChooser();
        chooser.setDialogTitle("設定のインポート");
        if (chooser.showOpenDialog(this) != JFileChooser.APPROVE_OPTION) {
            return;
        }
        File file = chooser.getSelectedFile();
        try {
            String text = Files.readString(file.toPath(), StandardCharsets.UTF_8);
            AppSettings settings = AppSettings.parse(text);
            if (!ruleModel.snapshot().isEmpty()) {
                int answer = JOptionPane.showConfirmDialog(
                        this,
                        "現在の置換ルールと設定を置き換えます。よろしいですか？",
                        "設定インポート",
                        JOptionPane.YES_NO_OPTION);
                if (answer != JOptionPane.YES_OPTION) {
                    return;
                }
            }
            applySettings(settings);
            remember(file);
            saveLastSessionQuietly();
            appendLog("設定を読み込みました: " + file.getAbsolutePath()
                    + "（ルール " + settings.getRules().size() + " 件）");
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "読み込みに失敗しました: " + e.getMessage(), "設定インポート", JOptionPane.ERROR_MESSAGE);
        }
    }

    private JFileChooser settingsChooser() {
        JFileChooser chooser = fileChooser(false);
        chooser.setFileFilter(new FileNameExtensionFilter("設定ファイル (*.txt, *.tsv)", "txt", "tsv"));
        chooser.setAcceptAllFileFilterUsed(true);
        return chooser;
    }

    private static File withSettingsExtension(File file) {
        String name = file.getName();
        if (name.contains(".")) {
            return file;
        }
        return new File(file.getParentFile(), name + ".txt");
    }

    private void runReplace() {
        stopEditing();
        List<Path> inputRoots = resolveInputRoots();
        if (inputRoots == null) {
            return;
        }
        ProcessOptions options = currentOptions();
        if (!options.isCells() && !options.isShapes() && !options.isComments()
                && !options.isHeadersFooters() && !options.isSheetNames()) {
            JOptionPane.showMessageDialog(this, "置換対象を 1 つ以上選んでください。", "入力エラー", JOptionPane.WARNING_MESSAGE);
            return;
        }
        List<ReplaceRule> rules = ruleModel.getRules();
        try {
            if (RichTextEngine.compile(rules, options.regexFlags()).isEmpty()) {
                JOptionPane.showMessageDialog(this, "有効な置換ルールがありません。検索パターンを入力してください。",
                        "入力エラー", JOptionPane.WARNING_MESSAGE);
                return;
            }
        } catch (IllegalArgumentException e) {
            JOptionPane.showMessageDialog(this, e.getMessage(), "正規表現エラー", JOptionPane.ERROR_MESSAGE);
            return;
        }

        boolean dumpText = dumpTextCheck.isSelected();
        setBusy(true);
        appendLog("---- 置換開始 " + TIME.format(LocalDateTime.now()) + " ----");
        new SwingWorker<ProcessResult, String>() {
            @Override
            protected ProcessResult doInBackground() throws Exception {
                List<Path> files = ExcelFiles.listExcelFiles(inputRoots, recursiveCheck.isSelected());
                if (files.isEmpty()) {
                    throw new IllegalArgumentException("Excel ファイルが見つかりません。");
                }
                publish("対象: " + files.size() + " 件");
                Path destDir = specifiedOutputDirectory();
                if (destDir != null) {
                    if (Files.exists(destDir) && !Files.isDirectory(destDir)) {
                        throw new IllegalArgumentException("出力先はフォルダを指定してください。\n" + destDir);
                    }
                    if (!Files.isDirectory(destDir)) {
                        Files.createDirectories(destDir);
                        publish("出力フォルダを作成: " + destDir.toAbsolutePath().normalize());
                    }
                }
                boolean overwrite = confirmOverwriteIfNeeded(files, inputRoots);
                ProcessResult total = new ProcessResult();
                for (Path file : files) {
                    Path output = resolveOutputFor(file, inputRoots, files);
                    if (Files.exists(output) && !overwrite) {
                        publish("スキップ（既存）: " + output);
                        continue;
                    }
                    publish("置換: " + file.getFileName() + " -> " + output);
                    Path beforeTxt = ExcelFiles.dumpPathBeside(file, output);
                    Path afterTxt = ExcelFiles.defaultDumpPath(output);
                    if (beforeTxt.equals(afterTxt)) {
                        beforeTxt = ExcelFiles.defaultDumpPath(file);
                    }
                    ExcelFiles.createParentDirectories(output);
                    if (dumpText) {
                        ExcelFiles.createParentDirectories(beforeTxt);
                        Files.writeString(beforeTxt, dumper.dumpFile(file), StandardCharsets.UTF_8);
                        publish("テキスト出力（変換前）: " + beforeTxt);
                    }
                    ProcessResult one = replacer.processFile(file, output, rules, options, this::publish);
                    total.merge(one);
                    if (dumpText) {
                        ExcelFiles.createParentDirectories(afterTxt);
                        Files.writeString(afterTxt, dumper.dumpFile(output), StandardCharsets.UTF_8);
                        publish("テキスト出力（変換後）: " + afterTxt);
                    }
                }
                return total;
            }

            @Override
            protected void process(List<String> chunks) {
                chunks.forEach(MainFrame.this::appendLog);
            }

            @Override
            protected void done() {
                setBusy(false);
                try {
                    ProcessResult result = get();
                    appendLog("完了: " + result.summary());
                    saveLastSessionQuietly();
                    JOptionPane.showMessageDialog(MainFrame.this, result.summary(), "完了", JOptionPane.INFORMATION_MESSAGE);
                } catch (Exception e) {
                    showError(e);
                }
            }
        }.execute();
    }

    private ProcessOptions currentOptions() {
        ProcessOptions options = new ProcessOptions();
        options.setCells(cellsCheck.isSelected());
        options.setShapes(shapesCheck.isSelected());
        options.setComments(commentsCheck.isSelected());
        options.setHeadersFooters(headersCheck.isSelected());
        options.setSheetNames(sheetNamesCheck.isSelected());
        options.setCaseInsensitive(caseCheck.isSelected());
        options.setMultiline(multilineCheck.isSelected());
        options.setRecolor(recolorCheck.isSelected());
        options.setReplacementColor(replacementColor);
        return options;
    }

    private List<Path> resolveInputRoots() {
        String text = inputField.getText().trim();
        if (text.isEmpty()) {
            JOptionPane.showMessageDialog(this, "入力ファイルまたはフォルダを指定してください。", "入力エラー", JOptionPane.WARNING_MESSAGE);
            return null;
        }
        List<Path> roots = new ArrayList<>();
        for (String part : text.split("\\s*;\\s*")) {
            if (part.isBlank()) {
                continue;
            }
            Path path = Path.of(part);
            if (!Files.exists(path)) {
                JOptionPane.showMessageDialog(this, "入力先が存在しません。\n" + part, "入力エラー", JOptionPane.WARNING_MESSAGE);
                return null;
            }
            roots.add(path);
        }
        if (roots.isEmpty()) {
            JOptionPane.showMessageDialog(this, "入力ファイルまたはフォルダを指定してください。", "入力エラー", JOptionPane.WARNING_MESSAGE);
            return null;
        }
        return roots;
    }

    private Path specifiedOutputDirectory() {
        String specified = outputField.getText().trim();
        return specified.isEmpty() ? null : Path.of(specified);
    }

    private Path resolveOutputFor(Path file, List<Path> inputRoots, List<Path> allFiles) {
        Path destDir = specifiedOutputDirectory();
        if (destDir == null) {
            return ExcelFiles.defaultReplacedPath(file);
        }
        return destDir.resolve(relativeOutputName(file, inputRoots, allFiles));
    }

    private static Path relativeOutputName(Path file, List<Path> inputRoots, List<Path> allFiles) {
        if (inputRoots.size() == 1 && Files.isDirectory(inputRoots.get(0))) {
            return inputRoots.get(0).relativize(file);
        }
        long sameName = allFiles.stream()
                .filter(path -> path.getFileName().equals(file.getFileName()))
                .count();
        if (sameName > 1 && file.getParent() != null) {
            return Path.of(file.getParent().getFileName().toString() + "_" + file.getFileName());
        }
        return Path.of(file.getFileName().toString());
    }

    private boolean confirmOverwriteIfNeeded(List<Path> files, List<Path> inputRoots) {
        long existing = files.stream()
                .map(file -> resolveOutputFor(file, inputRoots, files))
                .filter(Files::exists)
                .count();
        if (existing == 0) {
            return true;
        }
        boolean[] accepted = {false};
        String message = files.size() == 1
                ? resolveOutputFor(files.get(0), inputRoots, files) + "\nは既に存在します。上書きしますか？"
                : existing + " 件の出力ファイルが既に存在します。上書きしますか？（以降すべてに適用）";
        try {
            SwingUtilities.invokeAndWait(() -> accepted[0] = JOptionPane.showConfirmDialog(
                    this,
                    message,
                    "上書き確認",
                    JOptionPane.YES_NO_OPTION) == JOptionPane.YES_OPTION);
        } catch (Exception e) {
            return false;
        }
        return accepted[0];
    }

    private void chooseInputFiles() {
        JFileChooser chooser = fileChooser(false);
        chooser.setMultiSelectionEnabled(true);
        chooser.setFileFilter(excelFilter());
        chooser.setDialogTitle("Excel ファイルを選択（Ctrl で複数選択）");
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File[] selected = chooser.getSelectedFiles();
            if (selected.length == 0 && chooser.getSelectedFile() != null) {
                selected = new File[] {chooser.getSelectedFile()};
            }
            inputField.setText(joinFiles(selected));
            if (selected.length > 0) {
                remember(selected[0]);
            }
            saveLastSessionQuietly();
        }
    }

    private void chooseInputFolder() {
        JFileChooser chooser = fileChooser(true);
        chooser.setDialogTitle("入力フォルダを選択");
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            inputField.setText(chooser.getSelectedFile().getAbsolutePath());
            remember(chooser.getSelectedFile());
            saveLastSessionQuietly();
        }
    }

    private void chooseOutputFolder() {
        JFileChooser chooser = fileChooser(true);
        chooser.setDialogTitle("出力フォルダを選択");
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            outputField.setText(chooser.getSelectedFile().getAbsolutePath());
            remember(chooser.getSelectedFile());
            saveLastSessionQuietly();
        }
    }

    private JFileChooser fileChooser(boolean directories) {
        JFileChooser chooser = new JFileChooser();
        String last = prefs.get("lastDir", System.getProperty("user.home"));
        File dir = new File(last);
        if (dir.isDirectory()) {
            chooser.setCurrentDirectory(dir);
        }
        chooser.setFileSelectionMode(directories ? JFileChooser.DIRECTORIES_ONLY : JFileChooser.FILES_ONLY);
        return chooser;
    }

    private static FileNameExtensionFilter excelFilter() {
        return new FileNameExtensionFilter("Excel ファイル (*.xlsx, *.xlsm, *.xls)", "xlsx", "xlsm", "xls");
    }

    private void remember(File file) {
        File dir = file.isDirectory() ? file : file.getParentFile();
        if (dir != null) {
            prefs.put("lastDir", dir.getAbsolutePath());
        }
    }

    private void enableInputDrop() {
        new DropTarget(inputField, new DropTargetAdapter() {
            @Override
            public void drop(DropTargetDropEvent event) {
                try {
                    event.acceptDrop(DnDConstants.ACTION_COPY);
                    @SuppressWarnings("unchecked")
                    List<File> files = (List<File>) event.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
                    if (files == null || files.isEmpty()) {
                        return;
                    }
                    List<File> accepted = new ArrayList<>();
                    for (File file : files) {
                        if (file.isDirectory() || ExcelFiles.isExcelFile(file.toPath())) {
                            accepted.add(file);
                        }
                    }
                    if (accepted.isEmpty()) {
                        event.dropComplete(false);
                        return;
                    }
                    inputField.setText(joinFiles(accepted.toArray(File[]::new)));
                    remember(accepted.get(0));
                    saveLastSessionQuietly();
                    event.dropComplete(true);
                } catch (Exception ex) {
                    event.dropComplete(false);
                }
            }
        });
    }

    private void enableOutputDrop() {
        new DropTarget(outputField, new DropTargetAdapter() {
            @Override
            public void drop(DropTargetDropEvent event) {
                try {
                    event.acceptDrop(DnDConstants.ACTION_COPY);
                    @SuppressWarnings("unchecked")
                    List<File> files = (List<File>) event.getTransferable().getTransferData(DataFlavor.javaFileListFlavor);
                    if (files == null || files.isEmpty()) {
                        return;
                    }
                    File file = files.get(0);
                    File dir = file.isDirectory() ? file : file.getParentFile();
                    if (dir == null) {
                        event.dropComplete(false);
                        return;
                    }
                    outputField.setText(dir.getAbsolutePath());
                    remember(dir);
                    saveLastSessionQuietly();
                    event.dropComplete(true);
                } catch (Exception ex) {
                    event.dropComplete(false);
                }
            }
        });
    }

    private static String joinFiles(File[] files) {
        StringBuilder sb = new StringBuilder();
        for (File file : files) {
            if (sb.length() > 0) {
                sb.append("; ");
            }
            sb.append(file.getAbsolutePath());
        }
        return sb.toString();
    }

    private void setBusy(boolean busy) {
        replaceButton.setEnabled(!busy);
        dumpTextCheck.setEnabled(!busy);
        setCursor(busy ? Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR) : Cursor.getDefaultCursor());
    }

    private void stopEditing() {
        if (ruleTable.isEditing()) {
            ruleTable.getCellEditor().stopCellEditing();
        }
    }

    private void appendLog(String line) {
        logArea.append(line + System.lineSeparator());
        logArea.setCaretPosition(logArea.getDocument().getLength());
    }

    private void showError(Exception e) {
        Throwable cause = e;
        if (e instanceof java.util.concurrent.ExecutionException && e.getCause() != null) {
            cause = e.getCause();
        }
        String message = cause.getMessage() == null ? cause.getClass().getSimpleName() : cause.getMessage();
        if (cause instanceof EncryptedDocumentException) {
            message = "パスワード付きブックは未対応です。";
        }
        appendLog("エラー: " + message);
        JOptionPane.showMessageDialog(this, message, "エラー", JOptionPane.ERROR_MESSAGE);
    }

    private void chooseReplacementColor() {
        Color chosen = JColorChooser.showDialog(this, "置換後の文字色（全箇所共通）", replacementColor);
        if (chosen != null) {
            replacementColor = chosen;
            updateColorButton();
            saveLastSessionQuietly();
        }
    }

    private void updateColorButton() {
        colorButton.setBackground(replacementColor);
        colorButton.setForeground(contrast(replacementColor));
        colorButton.setText(String.format("#%06X", replacementColor.getRGB() & 0xFFFFFF));
        colorButton.setOpaque(true);
        colorButton.setEnabled(recolorCheck.isSelected());
    }

    private static JButton button(String text, ActionListener listener) {
        JButton button = new JButton(text);
        button.addActionListener(listener);
        return button;
    }

    private static Color contrast(Color color) {
        double luminance = (0.299 * color.getRed() + 0.587 * color.getGreen() + 0.114 * color.getBlue()) / 255;
        return luminance > 0.6 ? Color.BLACK : Color.WHITE;
    }
}
