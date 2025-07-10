package ru.unidubna;

import org.apache.poi.xwpf.usermodel.*;
import java.io.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class WordTemplateProcessor {
    private static final String ANSWER_PLACEHOLDER = "[ОТВЕТ]";
    private static final Pattern ANSWER_PATTERN = Pattern.compile("\\[ОТВЕТ\\]");


    public static void fillTemplate(String templatePath, String outputDirectory, List<Map<String, String>> allStudentsData) throws IOException {
        if (allStudentsData.isEmpty()) {
            System.out.println("Нет данных для обработки");
            return;
        }

        // if there's no template, we're creating it
        if (templatePath == null) {
            templatePath = outputDirectory + File.separator + "template_auto.docx";
            createTemplate(templatePath, allStudentsData);
            System.out.println("Создан автоматический шаблон: " + templatePath);
        }

        File outputDir = new File(outputDirectory);
        if (!outputDir.exists()) {
            outputDir.mkdirs();
        }

        // for each entry
        for (int i = 0; i < allStudentsData.size(); i++) {
            Map<String, String> studentData = allStudentsData.get(i);
            String studentName = studentData.getOrDefault("ФИО", "Студент_" + (i + 1));

            // file name gen
            String safeFileName = studentName.replaceAll("[^a-zA-Zа-яА-Я0-9\\s]", "").replaceAll("\\s+", "_");
            String outputPath = outputDirectory + File.separator + "справка_" + safeFileName + ".docx";

            fillSingleTemplate(templatePath, outputPath, studentData);
            System.out.println("Создан документ: " + outputPath + " для " + studentName);
        }
    }

    private static void fillSingleTemplate(String templatePath, String outputPath, Map<String, String> studentData) throws IOException {
        try (FileInputStream fis = new FileInputStream(templatePath);
             XWPFDocument document = new XWPFDocument(fis)) {

            for (Map.Entry<String, String> entry : studentData.entrySet()) {
                String question = entry.getKey();
                String answer = entry.getValue();

                if (answer == null || answer.trim().isEmpty()) {
                    answer = "Не указано";
                }

                replaceAnswerAfterQuestion(document, question, answer);
            }

            // Заменяем оставшиеся [ОТВЕТ] на "Не указано"
            replaceRemainingPlaceholders(document);

            try (FileOutputStream fos = new FileOutputStream(outputPath)) {
                document.write(fos);
            }
        }
    }

    private static void replaceAnswerAfterQuestion(XWPFDocument document, String question, String answer) {
        if (replaceInParagraphs(document.getParagraphs(), question, answer)) {
            return;
        }

        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    if (replaceInParagraphs(cell.getParagraphs(), question, answer)) {
                        return;
                    }
                }
            }
        }

        for (XWPFHeader header : document.getHeaderList()) {
            if (replaceInParagraphs(header.getParagraphs(), question, answer)) {
                return;
            }
        }

        for (XWPFFooter footer : document.getFooterList()) {
            if (replaceInParagraphs(footer.getParagraphs(), question, answer)) {
                return;
            }
        }
    }

    private static boolean replaceInParagraphs(List<XWPFParagraph> paragraphs, String question, String answer) {
        boolean questionFound = false;

        for (XWPFParagraph paragraph : paragraphs) {
            String paragraphText = paragraph.getText();
            if (paragraphText == null) continue;

            if (paragraphText.contains(question)) {
                questionFound = true;
                continue;
            }

            if (questionFound && paragraphText.contains(ANSWER_PLACEHOLDER)) {
                replacePlaceholderInParagraph(paragraph, ANSWER_PLACEHOLDER, answer);
                return true;
            }
        }

        return false;
    }

    private static void replaceRemainingPlaceholders(XWPFDocument document) {
        replaceRemainingInParagraphs(document.getParagraphs());

        for (XWPFTable table : document.getTables()) {
            for (XWPFTableRow row : table.getRows()) {
                for (XWPFTableCell cell : row.getTableCells()) {
                    replaceRemainingInParagraphs(cell.getParagraphs());
                }
            }
        }

        for (XWPFHeader header : document.getHeaderList()) {
            replaceRemainingInParagraphs(header.getParagraphs());
        }

        for (XWPFFooter footer : document.getFooterList()) {
            replaceRemainingInParagraphs(footer.getParagraphs());
        }
    }

    private static void replaceRemainingInParagraphs(List<XWPFParagraph> paragraphs) {
        for (XWPFParagraph paragraph : paragraphs) {
            String paragraphText = paragraph.getText();
            if (paragraphText != null && paragraphText.contains(ANSWER_PLACEHOLDER)) {
                replacePlaceholderInParagraph(paragraph, ANSWER_PLACEHOLDER, "Не указано");
            }
        }
    }

    private static void replacePlaceholderInParagraph(XWPFParagraph paragraph, String placeholder, String replacement) {
        List<XWPFRun> runs = paragraph.getRuns();

        StringBuilder fullText = new StringBuilder();
        for (XWPFRun run : runs) {
            String runText = run.getText(0);
            if (runText != null) {
                fullText.append(runText);
            }
        }

        String replacedText = fullText.toString().replace(placeholder, replacement);

        for (int i = runs.size() - 1; i >= 0; i--) {
            paragraph.removeRun(i);
        }

        XWPFRun newRun = paragraph.createRun();
        newRun.setText(replacedText);
    }

    private static void createTemplate(String templatePath, List<Map<String, String>> sampleData) throws IOException {
        try (XWPFDocument document = new XWPFDocument()) {

            XWPFParagraph title = document.createParagraph();
            title.setAlignment(ParagraphAlignment.CENTER);
            XWPFRun titleRun = title.createRun();
            titleRun.setText("СПРАВКА О ПРОХОЖДЕНИИ ОПРОСА");
            titleRun.setBold(true);
            titleRun.setFontSize(16);

            document.createParagraph();

            if (!sampleData.isEmpty()) {
                Map<String, String> sample = sampleData.get(0);

                for (String question : sample.keySet()) {
                    XWPFParagraph questionPara = document.createParagraph();
                    XWPFRun questionRun = questionPara.createRun();
                    questionRun.setText(question);
                    questionRun.setBold(true);

                    XWPFParagraph answerPara = document.createParagraph();
                    XWPFRun answerRun = answerPara.createRun();
                    answerRun.setText("Ответ: " + ANSWER_PLACEHOLDER);

                    document.createParagraph();
                }
            }

            try (FileOutputStream fos = new FileOutputStream(templatePath)) {
                document.write(fos);
            }
        }
    }
}