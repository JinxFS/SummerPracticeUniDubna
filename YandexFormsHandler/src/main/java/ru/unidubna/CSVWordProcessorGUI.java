package ru.unidubna;

import javax.swing.*;
import javax.swing.border.TitledBorder;
import javax.swing.filechooser.FileNameExtensionFilter;
import java.awt.*;
import java.io.File;
import java.io.IOException;
import com.opencsv.exceptions.CsvException;

public class CSVWordProcessorGUI extends JFrame {
    private JTextField csvFileField;
    private JTextField templateFileField;
    private JTextField outputDirField;
    private JButton csvBrowseButton;
    private JButton templateBrowseButton;
    private JButton outputBrowseButton;
    private JButton processButton;
    private JTextArea logArea;
    private JProgressBar progressBar;
    private JCheckBox useTemplateCheckbox;

    public CSVWordProcessorGUI() {
        initializeGUI();
    }

    private void initializeGUI() {
        setTitle("CSV to Word Template Processor");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(600, 500);
        setLocationRelativeTo(null);

        // main panel
        JPanel mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

        // file selection
        JPanel filePanel = createFileSelectionPanel();

        // log info output
        JPanel logPanel = createLogPanel();

        JPanel buttonPanel = createButtonPanel();

        mainPanel.add(filePanel, BorderLayout.NORTH);
        mainPanel.add(logPanel, BorderLayout.CENTER);
        mainPanel.add(buttonPanel, BorderLayout.SOUTH);

        add(mainPanel);

        updateTemplateFieldState();
    }

    private JPanel createFileSelectionPanel() {
        JPanel panel = new JPanel(new GridBagLayout());
        panel.setBorder(new TitledBorder("Настройки файлов"));
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);

        gbc.gridx = 0; gbc.gridy = 0; gbc.anchor = GridBagConstraints.WEST;
        panel.add(new JLabel("CSV файл:"), gbc);

        gbc.gridx = 1; gbc.fill = GridBagConstraints.HORIZONTAL; gbc.weightx = 1.0;
        csvFileField = new JTextField();
        csvFileField.setEditable(false);
        panel.add(csvFileField, gbc);

        gbc.gridx = 2; gbc.fill = GridBagConstraints.NONE; gbc.weightx = 0;
        csvBrowseButton = new JButton("Обзор...");
        csvBrowseButton.addActionListener(e -> browseCsvFile());
        panel.add(csvBrowseButton, gbc);

        // if using template (check/uncheck)
        gbc.gridx = 0; gbc.gridy = 1; gbc.gridwidth = 3;
        useTemplateCheckbox = new JCheckBox("Использовать готовый шаблон Word");
        useTemplateCheckbox.addActionListener(e -> updateTemplateFieldState());
        panel.add(useTemplateCheckbox, gbc);

        gbc.gridx = 0; gbc.gridy = 2; gbc.gridwidth = 1; gbc.anchor = GridBagConstraints.WEST;
        panel.add(new JLabel("Шаблон Word:"), gbc);

        gbc.gridx = 1; gbc.fill = GridBagConstraints.HORIZONTAL; gbc.weightx = 1.0;
        templateFileField = new JTextField();
        templateFileField.setEditable(false);
        panel.add(templateFileField, gbc);

        gbc.gridx = 2; gbc.fill = GridBagConstraints.NONE; gbc.weightx = 0;
        templateBrowseButton = new JButton("Обзор...");
        templateBrowseButton.addActionListener(e -> browseTemplateFile());
        panel.add(templateBrowseButton, gbc);

        // save path
        gbc.gridx = 0; gbc.gridy = 3; gbc.anchor = GridBagConstraints.WEST;
        panel.add(new JLabel("Папка для сохранения:"), gbc);

        gbc.gridx = 1; gbc.fill = GridBagConstraints.HORIZONTAL; gbc.weightx = 1.0;
        outputDirField = new JTextField();
        outputDirField.setEditable(false);
        panel.add(outputDirField, gbc);

        gbc.gridx = 2; gbc.fill = GridBagConstraints.NONE; gbc.weightx = 0;
        outputBrowseButton = new JButton("Обзор...");
        outputBrowseButton.addActionListener(e -> browseOutputDirectory());
        panel.add(outputBrowseButton, gbc);

        return panel;
    }

    private JPanel createLogPanel() {
        JPanel panel = new JPanel(new BorderLayout());
        panel.setBorder(new TitledBorder("Лог выполнения"));

        logArea = new JTextArea(10, 50);
        logArea.setEditable(false);
        logArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        JScrollPane scrollPane = new JScrollPane(logArea);

        panel.add(scrollPane, BorderLayout.CENTER);

        return panel;
    }

    private JPanel createButtonPanel() {
        JPanel panel = new JPanel(new FlowLayout());

        processButton = new JButton("Обработать");
        processButton.addActionListener(e -> processFiles());
        processButton.setPreferredSize(new Dimension(120, 30));

        JButton clearLogButton = new JButton("Очистить лог");
        clearLogButton.addActionListener(e -> logArea.setText(""));
        clearLogButton.setPreferredSize(new Dimension(120, 30));

        progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        progressBar.setPreferredSize(new Dimension(200, 25));

        panel.add(processButton);
        panel.add(clearLogButton);
        panel.add(progressBar);

        return panel;
    }

    private void updateTemplateFieldState() {
        boolean useTemplate = useTemplateCheckbox.isSelected();
        templateFileField.setEnabled(useTemplate);
        templateBrowseButton.setEnabled(useTemplate);

        if (!useTemplate) {
            templateFileField.setText("");
        }
    }

    private void browseCsvFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("CSV файлы", "csv"));

        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            csvFileField.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void browseTemplateFile() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Word документы", "docx", "doc"));

        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            templateFileField.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void browseOutputDirectory() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        if (fileChooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            outputDirField.setText(fileChooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void processFiles() {
        // obligatory fields check
        if (csvFileField.getText().trim().isEmpty()) {
            showError("Необходимо выбрать CSV файл");
            return;
        }

        if (outputDirField.getText().trim().isEmpty()) {
            showError("Необходимо выбрать папку для сохранения");
            return;
        }

        if (useTemplateCheckbox.isSelected() && templateFileField.getText().trim().isEmpty()) {
            showError("Необходимо выбрать файл шаблона Word");
            return;
        }

        // Processing in da dedicated thread
        SwingWorker<Void, String> worker = new SwingWorker<Void, String>() {
            @Override
            protected Void doInBackground() throws Exception {
                processFilesInBackground();
                return null;
            }

            @Override
            protected void process(java.util.List<String> chunks) {
                for (String message : chunks) {
                    logArea.append(message + "\n");
                    logArea.setCaretPosition(logArea.getDocument().getLength());
                }
            }

            @Override
            protected void done() {
                processButton.setEnabled(true);
                progressBar.setValue(0);
                progressBar.setString("Готово");

                try {
                    get(); // exception check
                    publish("✓ Обработка завершена успешно!");
                } catch (Exception e) {
                    publish("✗ Ошибка: " + e.getMessage());
                    e.printStackTrace();
                }
            }
        };

        processButton.setEnabled(false);
        progressBar.setIndeterminate(true);
        progressBar.setString("Обработка...");
        logArea.setText("");

        worker.execute();
    }

    private void processFilesInBackground() throws IOException, CsvException {
        String csvPath = csvFileField.getText().trim();
        String templatePath = useTemplateCheckbox.isSelected() ? templateFileField.getText().trim() : null;
        String outputPath = outputDirField.getText().trim();

        if (!new File(csvPath).exists()) {
            throw new IOException("CSV файл не найден: " + csvPath);
        }

        if (templatePath != null && !new File(templatePath).exists()) {
            throw new IOException("Файл шаблона не найден: " + templatePath);
        }

        SwingUtilities.invokeLater(() -> {
            logArea.append("Начинаем обработку...\n");
            logArea.append("CSV файл: " + csvPath + "\n");
            logArea.append("Шаблон: " + (templatePath != null ? templatePath : "автоматический") + "\n");
            logArea.append("Папка сохранения: " + outputPath + "\n");
            logArea.append("------------------------\n");
        });

        SwingUtilities.invokeLater(() -> logArea.append("Чтение CSV файла...\n"));
        CSVDataFetcher csvDataFetcher = new CSVDataFetcher(csvPath);

        SwingUtilities.invokeLater(() -> logArea.append("Создание Word документов...\n"));
        WordTemplateProcessor.fillTemplate(templatePath, outputPath, csvDataFetcher.getCsvData());

        SwingUtilities.invokeLater(() -> logArea.append("Обработка завершена!\n"));
    }

    private void showError(String message) {
        JOptionPane.showMessageDialog(this, message, "Ошибка", JOptionPane.ERROR_MESSAGE);
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            try {
                UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
            } catch (Exception e) {
                e.printStackTrace();
            }

            new CSVWordProcessorGUI().setVisible(true);
        });
    }
}