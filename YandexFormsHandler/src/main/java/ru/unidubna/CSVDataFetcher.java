package ru.unidubna;
import java.io.*;
import java.util.*;
import java.util.regex.Pattern;
import com.opencsv.CSVReader;
import com.opencsv.exceptions.CsvException;

public class CSVDataFetcher {
    private List<Map<String, String>> csvData = new ArrayList<>();

    public CSVDataFetcher(String csvPath) throws IOException, CsvException {
        try (CSVReader reader = new CSVReader(new FileReader(csvPath))) {
            List<String[]> allRows = reader.readAll();
            if (allRows.isEmpty()) return;

            String[] headers = allRows.get(0);
            List<String[]> dataRows = allRows.subList(1, allRows.size());

            // Model for the data retrieved from .csv file
            Map<String, List<String>> questionToColumns = groupQuestions(headers);

            for (String[] row : dataRows) {
                Map<String, String> rowMap = new LinkedHashMap<>();
                Set<String> processedQuestions = new HashSet<>();

                // Handling every question for every row
                for (Map.Entry<String, List<String>> entry : questionToColumns.entrySet()) {
                    String mainQuestion = entry.getKey();
                    List<String> columns = entry.getValue();

                    StringBuilder answerBuilder = new StringBuilder();
                    String pointsValue = "";
                    boolean hasVariants = false;

                    int mainColIndex = Arrays.asList(headers).indexOf(mainQuestion);
                    if (mainColIndex != -1 && mainColIndex < row.length && !row[mainColIndex].isEmpty()) {
                        answerBuilder.append(row[mainColIndex]);
                        hasVariants = true;
                    }

                    for (String column : columns) {
                        int colIndex = Arrays.asList(headers).indexOf(column);
                        if (colIndex < row.length && !row[colIndex].isEmpty()) {
                            String[] parts = column.split(Pattern.quote(" / "));
                            if (parts.length > 1) {
                                if (parts[1].equals("Баллы")) {
                                    pointsValue = row[colIndex];
                                } else {
                                    // Append variant to answer
                                    if (answerBuilder.length() > 0) {
                                        answerBuilder.append(", ");
                                    }
                                    answerBuilder.append(parts[1]);
                                    hasVariants = true;
                                }
                            }
                        }
                    }

                    String finalAnswer = answerBuilder.toString();
                    if (!pointsValue.isEmpty()) {
                        if (!finalAnswer.isEmpty()) {
                            finalAnswer += "; баллы - " + pointsValue;
                        } else {
                            // Only points, no variants
                            finalAnswer = "баллы - " + pointsValue;
                        }
                    }

                    rowMap.put(mainQuestion, finalAnswer);
                    processedQuestions.add(mainQuestion);
                }

                // For fields without " / " in the header that weren't processed above
                for (int i = 0; i < headers.length; i++) {
                    String header = headers[i];
                    if (!header.contains(" / ") && !processedQuestions.contains(header)) {
                        if (i < row.length) {
                            rowMap.put(header, row[i]);
                        }
                    }
                }

                csvData.add(rowMap);
            }

            // Test output
            for (Map<String, String> tmpMap : csvData) {
                System.out.println(tmpMap);
            }
        }
    }

    // Group by main question (to handle "<question> / <option or score>")
    private Map<String, List<String>> groupQuestions(String[] headers) {
        Map<String, List<String>> questionToColumns = new LinkedHashMap<>();

        for (String header : headers) {
            if (header.contains(" / ")) {
                String[] parts = header.split(Pattern.quote(" / "));
                String mainQuestion = parts[0];
                questionToColumns.computeIfAbsent(mainQuestion, k -> new ArrayList<>()).add(header);
            }
        }

        return questionToColumns;
    }

    public List<Map<String, String>> getCsvData() {
        return csvData;
    }
}