package ddbmerge;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import java.lang.reflect.Type;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
// import java.text.SimpleDateFormat;
// import java.util.Date;

public class Main {
    public static Map<String, Object> dataMap;

    public static void main(String[] args) {
        String inputFilePath = "LTR- Disc Req to Client.docx";
        String outputFilePath = "output_merged.docx";
        dataMap = new HashMap<>();
        String json = "{" +
                "  \"filePath\": {" +
                "    \"Object\": \"Matter\"," +
                "    \"Templates\": [" +
                "      \"TestDocgen.docx\"" +
                "    ]" +
                "  }," +
                "  \"fileOutput\": \"docx\"," +
                "  \"data\": {" +
                "    \"Matter\": {" +
                "      \"Status\": \"Open\"," +
                "       \"Client\": {\"Birthday\": \"2024-01-01\"}," +
                "      \"Handling_attorney\": {" +
                "        \"Name\": \"Handling ATORNEY\"," +
                "        \"signature\": \"https://example.com\"" +
                "      }," +
                "      \"Managing_Attorney__r\": {" +
                "        \"Name\": \"Hayden Dudley\"," +
                "        \"Phone\": \"225-907-9252\"," +
                "        \"Street\": \"1075 Government St\"," +
                "        \"City\": \"Baton Rouge\"," +
                "        \"State\": \"Louisiana\"," +
                "        \"PostalCode\": \"70830\"," +
                "        \"Fax\": \"225-444-4444\"," +
                "        \"Email\": \"hdudley@dudleydebosier.com\"" +
                "      }," +
                "      \"litify_pm__lit_Case_Manager__r\": {" +
                "        \"Name\": \"Michael Dudley\"," +
                "        \"Phone\": \"225-907-9156\"," +
                "        \"Title\": \"Paralegal\"," +
                "        \"Fax\": \"225-343-3333\"," +
                "        \"Email\": \"mdudley@dudleydebosier.com\"" +
                "      }" +
                "    }," +
                " \"Today\": \"2025-05-04\"," +
                "    \"Medical_Provider\": {" +
                "      \"Name\": \"Ochsner Health\"" +
                "    }," +
                "    \"Roles\": [" +
                "      {" +
                "        \"Name\": \"John Doe\"," +
                "        \"Type\": \"Witness\"," +
                "        \"Notes\": \"Notes 1\"" +
                "      }," +
                "      {" +
                "        \"Name\": \"Jane Doe\"," +
                "        \"Type\": \"Attorney\"," +
                "        \"Notes\": \"Notes 2\"" +
                "      }" +
                "    ]," +
                "    \"fullAndFinal\": true," +
                "    \"Louisiana\": false" +
                "  }" +
                "}";
        Gson gson = new Gson();
        Type type = new TypeToken<Map<String, Object>>() {
        }.getType();
        Map<String, Object> jsonMap = gson.fromJson(json, type);
        dataMap = (Map<String, Object>) jsonMap.get("data");
        try {
            XWPFDocument document = readDocxFile(inputFilePath);

            if (document != null) {

                // Replace tag in headers
                for (XWPFHeader header : document.getHeaderList()) {
                    // process Header tables
                    for (XWPFTable table : header.getTables()) {
                        processTable(table);
                    }
                    // process Header paragraphs
                    for (XWPFParagraph paragraph : header.getParagraphs()) {
                        mergeTagInParagraph(paragraph, dataMap);
                    }
                }

                // Replace tag in main document paragraphs
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    mergeTagInParagraph(paragraph, dataMap);
                }

                // process Document tables
                for (XWPFTable table : document.getTables()) {
                    processTable(table);
                }

                // Replace tag in footers
                for (XWPFFooter footer : document.getFooterList()) {
                    for (XWPFParagraph paragraph : footer.getParagraphs()) {
                        mergeTagInParagraph(paragraph, dataMap);
                    }
                    // process Footer tables
                    for (XWPFTable table : footer.getTables()) {
                        processTable(table);
                    }
                }

                generateDocxFile(document, outputFilePath);
                System.out.println("Successfully processed document, headers, and footers.");

            } else {
                System.err.println("Error reading the DOCX file.");
            }

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static XWPFDocument readDocxFile(String filePath) throws IOException {
        try (FileInputStream fis = new FileInputStream(filePath)) {
            return new XWPFDocument(fis);
        } catch (IOException e) {
            System.err.println("Error opening or reading the DOCX file: " + e.getMessage());
            throw e;
        }
    }

    public static void generateDocxFile(XWPFDocument document, String filePath) throws IOException {
        try (FileOutputStream fos = new FileOutputStream(filePath)) {
            document.write(fos);
        } catch (IOException e) {
            System.err.println("Error writing to the DOCX file: " + e.getMessage());
            throw e;
        } finally {
            document.close();
        }
    }

    private static void processTable(XWPFTable table) {
        List<XWPFTableRow> tableRows = new ArrayList<>(table.getRows());
        Map<Integer, List<XWPFTableRow>> tableRowsToAdd = new HashMap<>();
        List<Integer> rowsToDelete = new ArrayList<>();

        if (tableRows != null) {
            int rowIndex = 0;
            for (XWPFTableRow tableRow : tableRows) {
                List<XWPFTableCell> tableCells = tableRow.getTableCells();
                boolean replicateRow = false;
                if (tableCells != null && !tableCells.isEmpty()) {
                    XWPFTableCell firstCell = tableCells.get(0);
                    List<XWPFParagraph> firstCellParagraphs = firstCell.getParagraphs();
                    if (firstCellParagraphs != null && !firstCellParagraphs.isEmpty()) {
                        XWPFParagraph firstCellParagraph = firstCellParagraphs.get(0);
                        String paragraphText = firstCellParagraph.getText();
                        Pattern pattern = Pattern.compile("\\{\\{startRow\\.(.*?)\\}\\}");
                        Matcher matcher = pattern.matcher(paragraphText);

                        List<Runnable> deferredReplacements = new ArrayList<>();

                        while (matcher.find()) {
                            replicateRow = true;
                            String startRowKey = matcher.group(1);
                            if (startRowKey != null) {
                                List<XWPFTableRow> rowsToAdd = new ArrayList<>();
                                List<Map<String, Object>> repeatMap = (List<Map<String, Object>>) getNestedValueList(
                                        dataMap,
                                        Arrays.asList(startRowKey.split("\\.")), 0);

                                if (repeatMap != null) {
                                    for (Map<String, Object> listMap : repeatMap) {
                                        org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow ctRow = null;
                                        try {
                                            ctRow = org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow.Factory
                                                    .parse(tableRow.getCtRow().newInputStream());
                                        } catch (Exception e) {
                                            System.err.println("Error converting to ctrow: " + e.getMessage());
                                        }
                                        XWPFTableRow newTableRow = new XWPFTableRow(ctRow, table);
                                        for (XWPFTableCell replicateTableCell : newTableRow.getTableCells()) {
                                            if (replicateTableCell != null) {
                                                for (XWPFParagraph cellParagraph : replicateTableCell.getParagraphs()) {
                                                    mergeTagInParagraph(cellParagraph, listMap);
                                                }
                                            }
                                        }
                                        rowsToAdd.add(newTableRow);
                                    }
                                    tableRowsToAdd.put(rowIndex, rowsToAdd);
                                }
                            }

                            // Create a deferred action to remove the tag
                            int start = matcher.start();
                            int end = matcher.end();
                            deferredReplacements.add(() -> replaceTextRange(firstCellParagraph, start, end, ""));
                        }

                        // Execute deferred replacements *after* the loop to avoid index issues
                        for (Runnable replacement : deferredReplacements) {
                            replacement.run();
                        }
                    }
                }

                if (replicateRow) {
                    rowsToDelete.add(rowIndex);
                }
                rowIndex++;
            }

            // Delete original rows
            for (int i = rowsToDelete.size() - 1; i >= 0; i--) {
                table.removeRow(rowsToDelete.get(i));
            }

            // Add new rows
            for (Integer indexInteger : tableRowsToAdd.keySet()) {
                List<XWPFTableRow> rowsToAdd = tableRowsToAdd.get(indexInteger);
                if (rowsToAdd != null) {
                    for (XWPFTableRow newRow : rowsToAdd) {
                        try {
                            table.addRow(newRow, indexInteger);
                        } catch (Exception e) {
                            System.err.println("Error adding row: " + e);
                        }
                    }
                }
            }

            // Re-process the table for other merge tags in the newly added rows
            for (XWPFTableRow tableRow : table.getRows()) {
                for (XWPFTableCell tableCell : tableRow.getTableCells()) {
                    for (XWPFParagraph cellParagraph : tableCell.getParagraphs()) {
                        mergeTagInParagraph(cellParagraph, dataMap);
                    }
                }
            }
        }
    }

    private static void replaceTextRange(XWPFParagraph paragraph, int start, int end, String replacement) {
        String paragraphText = paragraph.getText();
        if (start >= 0 && end <= paragraphText.length() && start <= end) {
            int currentPos = 0;
            List<XWPFRun> runsToRemove = new ArrayList<>();
            List<XWPFRun> runs = paragraph.getRuns();

            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = run.getText(0) == null ? "" : run.getText(0);
                int runStart = currentPos;
                int runEnd = currentPos + runText.length();

                if (runEnd > start && runStart < end) {
                    runsToRemove.add(run);
                }
                currentPos = runEnd;
            }

            // Remove the identified runs
            for (XWPFRun run : runsToRemove) {
                paragraph.removeRun(paragraph.getRuns().indexOf(run));
            }

            // Insert the replacement text if needed
            if (!replacement.isEmpty()) {
                XWPFRun newRun = paragraph.createRun();
                newRun.setText(replacement);
            }

            // Re-set the paragraph text to reflect the changes (important for internal
            // state)
            StringBuilder sb = new StringBuilder();
            for (XWPFRun run : paragraph.getRuns()) {
                sb.append(run.getText(0) == null ? "" : run.getText(0));
            }
            paragraph.getCTP().addNewR(); // Trigger a re-evaluation of the text
        }
    }

    private static void copyRunFormatting(XWPFRun sourceRun, XWPFRun targetRun) {
        targetRun.setFontFamily(sourceRun.getFontFamily());
        if (sourceRun.getFontSizeAsDouble() != null) {
            targetRun.setFontSize(sourceRun.getFontSizeAsDouble());
        }
        targetRun.setBold(sourceRun.isBold());
        targetRun.setItalic(sourceRun.isItalic());
        targetRun.setUnderline(sourceRun.getUnderline());
        targetRun.setColor(sourceRun.getColor());
        targetRun.setStrikeThrough(sourceRun.isStrikeThrough());
        targetRun.setTextPosition(sourceRun.getTextPosition());
        targetRun.setStyle(sourceRun.getStyle());

    }

    public static void mergeTagInParagraph(XWPFParagraph paragraph, Map<String, Object> tagMap) {
        List<XWPFRun> runs = paragraph.getRuns();
        if (runs == null || runs.isEmpty())
            return;

        StringBuilder fullText = new StringBuilder();
        List<RunInfo> runInfoList = new ArrayList<>();
        int currentPos = 0;

        for (int i = 0; i < runs.size(); i++) {
            String text = runs.get(i).getText(0);
            if (text == null)
                continue;

            fullText.append(text);
            runInfoList.add(new RunInfo(i, currentPos, currentPos + text.length(), text));
            currentPos += text.length();
        }

        String paragraphStr = fullText.toString();
        Pattern pattern = Pattern.compile("\\{\\{(.*?)\\}\\}");
        Matcher matcher = pattern.matcher(paragraphStr);

        List<TagMatch> matches = new ArrayList<>();

        while (matcher.find()) {
            String tag = matcher.group(1);
            int tagStart = matcher.start();
            int tagEnd = matcher.end();

            int startRun = -1, endRun = -1;
            for (RunInfo info : runInfoList) {
                if (startRun == -1 && tagStart >= info.start && tagStart < info.end) {
                    startRun = info.index;
                }
                if (tagEnd > info.start && tagEnd <= info.end) {
                    endRun = info.index;
                }
            }

            if (startRun != -1 && endRun != -1) {
                matches.add(new TagMatch(tag, tagStart, tagEnd, startRun, endRun));
            }
        }

        // Process in reverse to safely remove runs
        Collections.reverse(matches);

        for (TagMatch match : matches) {
            Object value = getNestedValue(tagMap, match.tag);
            String replacement = value != null ? value.toString() : "";

            // Update first run's text
            XWPFRun firstRun = paragraph.getRuns().get(match.startRun);
            String originalText = firstRun.getText(0);
            int startOffset = match.startChar - runInfoList.get(match.startRun).start;
            String prefix = originalText.substring(0, startOffset);
            String suffix = "";

            if (match.startRun == match.endRun) {
                int endOffset = match.endChar - runInfoList.get(match.endRun).start;
                suffix = originalText.substring(endOffset);
            } else {
                // Tag ends in another run — handle suffix in that run
                XWPFRun endRun = paragraph.getRuns().get(match.endRun);
                String endText = endRun.getText(0);
                int endOffset = match.endChar - runInfoList.get(match.endRun).start;
                suffix = endText.substring(endOffset);
            }

            firstRun.setText(prefix + replacement + suffix, 0);

            // Remove intermediate runs
            for (int i = match.endRun; i > match.startRun; i--) {
                paragraph.removeRun(i);
            }
        }
    }

    // Process standard tags
    public static Object getNestedValue(Map<String, Object> data, String path) {
        List<String> keys = Arrays.asList(path.split("\\."));
        if (keys.contains("startRow")) {
            return "";
        } else {
            return getNestedValueRecursive(data, keys, 0, path);

        }
    }

    private static Object getNestedValueRecursive(Map<String, Object> currentLevel, List<String> keys, int index,
            String path) {
        if (currentLevel == null || index >= keys.size()) {
            System.err.println("Invalid path or null map at index: " + index);
            return null; // Return null if the path is invalid
        }

        String currentKey = keys.get(index);
        Object value = currentLevel.get(currentKey);

        if (value == null) {
            System.err.println("Key not found: " + currentKey);
            return "{{" + path + "}}"; // Key not found
        }

        if (index == keys.size() - 1) {
            return value;

        } else if (value instanceof Map) {
            // If the value is a map, continue recursively
            return getNestedValueRecursive((Map<String, Object>) value, keys, index + 1, path);
        } else if (value instanceof List) {
            // If the value is a list, get the first element and continue if it's a map
            List<?> list = (List<?>) value;
            if (!list.isEmpty() && list.get(0) instanceof Map) {
                return getNestedValueRecursive((Map<String, Object>) list.get(0), keys, index
                        + 1, path);
            } else {
                System.err.println("List is empty or does not contain a map at key: " +
                        currentKey);
                return null;
            }
        } else {
            System.err.println("Intermediate value is not a map or list for key: " +
                    currentKey);
            return null; // Path not found or intermediate level is not a map or list
        }
    }

    private static Object getNestedValueList(Map<String, Object> currentLevel, List<String> keys, int index) {
        String currentKey = keys.get(index);
        Object value = currentLevel.get(currentKey);
        if (value == null) {
            return null;
        }
        if (index == keys.size() - 1) {
            return value;
        } else if (value instanceof String) {
            return getNestedValueList((Map<String, Object>) value, keys, index + 1);
        } else {
            return null;
        }

    }

    static class RunInfo {
        int index, start, end;
        String text;

        RunInfo(int index, int start, int end, String text) {
            this.index = index;
            this.start = start;
            this.end = end;
            this.text = text;
        }
    }

    static class TagMatch {
        String tag;
        int startChar, endChar;
        int startRun, endRun;

        TagMatch(String tag, int startChar, int endChar, int startRun, int endRun) {
            this.tag = tag;
            this.startChar = startChar;
            this.endChar = endChar;
            this.startRun = startRun;
            this.endRun = endRun;
        }
    }

}
