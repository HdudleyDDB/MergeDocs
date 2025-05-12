package ddbmerge;

import org.apache.poi.xwpf.usermodel.*;

import com.google.gson.Gson;
import com.google.gson.reflect.TypeToken;
import java.lang.reflect.Type;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Arrays;
import java.util.HashMap;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.text.SimpleDateFormat;
import java.util.Date;

public class ProcessHeadersFooters {
    public static Map<String, Object> dataMap;

    public static void main(String[] args) {
        String inputFilePath = "TestMerge.docx";
        String outputFilePath = "output_merged.docx";
        String tagToReplace = "{{TestMerge}}";
        String replacementText = "Hello World";
        String headerReplacementText = "Header Text";
        String footerReplacementText = "Footer Text";
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
                    for (XWPFParagraph paragraph : header.getParagraphs()) {
                        mergeTagInParagraph(paragraph);
                    }
                }

                // Replace tag in main document paragraphs
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    mergeTagInParagraph(paragraph);
                }

                // Replace tag in footers
                for (XWPFFooter footer : document.getFooterList()) {
                    for (XWPFParagraph paragraph : footer.getParagraphs()) {
                        mergeTagInParagraph(paragraph);
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

    public static void mergeTagInParagraph(XWPFParagraph paragraph) {
        List<XWPFRun> runs = paragraph.getRuns();
        StringBuilder paragraphText = new StringBuilder();
        int tagStartIndex = -1;
        int tagEndIndex = -1;
        int firstTagRunIndex = -1;
        int lastTagRunIndex = -1;
        Pattern pattern = Pattern.compile("\\{\\{(.*?)\\}\\}"); // Correct Pattern
        Matcher matcher = null;

        // Build the complete paragraph text first. This is crucial for correct
        // matching.
        for (XWPFRun run : runs) {
            String runText = run.getText(0) == null ? "" : run.getText(0);
            paragraphText.append(runText);
        }
        String paragraphTextStr = paragraphText.toString(); // Store as String for efficiency

        matcher = pattern.matcher(paragraphTextStr); // Create matcher with full text

        // Iterate through each match found in the paragraph
        while (matcher.find()) {
            String tagWithBraces = matcher.group(0);
            String tagBody = matcher.group(1);
            String replacement = (String) getNestedValue(dataMap, tagBody); // Get replacement value.

            tagStartIndex = matcher.start();
            tagEndIndex = matcher.end();

            // Find the runs that contain the tag. This logic is now correct.
            firstTagRunIndex = 0;
            int currentRunLength = 0;
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = run.getText(0) == null ? "" : run.getText(0);
                int runTextLength = runText.length();

                if (tagStartIndex >= currentRunLength && tagStartIndex < currentRunLength + runTextLength) {
                    firstTagRunIndex = i;
                    break; // Found the first run
                }
                currentRunLength += runTextLength;
            }

            lastTagRunIndex = runs.size() - 1;
            currentRunLength = paragraphTextStr.length();
            for (int i = runs.size() - 1; i >= 0; i--) {
                XWPFRun run = runs.get(i);
                String runText = run.getText(0) == null ? "" : run.getText(0);
                int runTextLength = runText.length();
                currentRunLength -= runTextLength;
                if (tagEndIndex > currentRunLength && tagEndIndex <= currentRunLength + runTextLength) {
                    lastTagRunIndex = i;
                    break;
                }
            }
            // Store formatting of the first run.
            XWPFRun formattingSourceRun = runs.get(firstTagRunIndex);

            // Create a new run for the replacement.
            XWPFRun replacementRun = paragraph.createRun();
            copyRunFormatting(formattingSourceRun, replacementRun); // Copy formatting
            replacementRun.setText(replacement);

            // Remove the runs containing the tag, in reverse order to avoid index issues.
            for (int i = lastTagRunIndex; i >= firstTagRunIndex; i--) {
                paragraph.removeRun(i);
            }
            // Insert the replacement run.
            paragraph.insertNewRun(firstTagRunIndex);
            runs = paragraph.getRuns();
            try {
                runs.set(firstTagRunIndex, replacementRun);

            } catch (Exception e) {
                // TODO: handle exception
                System.err.println("error setting replacement run " + e.getMessage());
            }

        }
    }

    // Process standard tags
    public static Object getNestedValue(Map<String, Object> data, String path) {
        List<String> keys = Arrays.asList(path.split("\\."));
        return getNestedValueRecursive(data, keys, 0, path);
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
}
