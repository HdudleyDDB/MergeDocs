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

public class Main {
    public static Map<String, Object> dataMap;
    public static void main(String[] args) {
        String inputFilePath = "MergeDocs\\TestMerge.docx";
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
                "       \"Client\": {\"Birthday\": \"2024-01-01\"},"+
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
            Type type = new TypeToken<Map<String, Object>>() {}.getType();
            Map<String, Object> jsonMap = gson.fromJson(json, type);
            dataMap = (Map<String, Object>) jsonMap.get("data");
        try {
            XWPFDocument document = readDocxFile(inputFilePath);

            if (document != null) {


                // Replace tag in headers
                for (XWPFHeader header : document.getHeaderList()) {
                    for (XWPFParagraph paragraph : header.getParagraphs()) {
                        mergeTagInParagraph(paragraph, tagToReplace, headerReplacementText);
                    }
                }

                // Replace tag in main document paragraphs
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    mergeTagInParagraph(paragraph, tagToReplace, replacementText);
                }

                // Replace tag in footers
                for (XWPFFooter footer : document.getFooterList()) {
                    for (XWPFParagraph paragraph : footer.getParagraphs()) {
                        mergeTagInParagraph(paragraph, tagToReplace, footerReplacementText);
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

    public static void mergeTagInParagraph(XWPFParagraph paragraph, String tag, String replacement) {
        List<XWPFRun> runs = paragraph.getRuns();
        StringBuilder paragraphText = new StringBuilder();
        int tagStartIndex = -1;
        int tagEndIndex = -1;
        int firstTagRunIndex = -1;
        int lastTagRunIndex = -1;

        //initialize pattern and matcher

        // First pass: Find the tag and the indices of the runs it spans
        for (int i = 0; i < runs.size(); i++) {
            XWPFRun run = runs.get(i);
            String runText = run.getText(0) == null ? "" : run.getText(0);
            paragraphText.append(runText);
            int endIndexInParagraph = paragraphText.length();
            Pattern pattern = Pattern.compile("\\{\\{(.*?)\\}\\}");

            Matcher matcher = pattern.matcher(paragraphText);

            //find paragraph text
            while(matcher.find()){
                String tagWithBraces = matcher.group(0);

                if (tagStartIndex == -1 && paragraphText.toString().contains(tagWithBraces)) {
                    tagStartIndex = paragraphText.indexOf(tagWithBraces);
                    firstTagRunIndex = i;
                    System.out.println("Tag Start Index" + tagStartIndex);
                    System.out.println("Found Tag: "+ tagWithBraces);
                    System.out.println("Tag Name:" + matcher.group(1));
                }
                if (tagStartIndex != -1 && endIndexInParagraph >= tagStartIndex + tagWithBraces.length() && lastTagRunIndex == -1) {
                    lastTagRunIndex = i;
                    tagEndIndex = tagStartIndex + tagWithBraces.length();
                }

                //replace tag
                String tagBody = matcher.group(1);
                String newReplacement = (String) getNestedValue(dataMap, tagBody);
                if (tagStartIndex != -1) {
                    System.err.println("Start Merge Replacement");
                    // Store the formatting of the *first* run of the tag
                    XWPFRun formattingSourceRun = null;
                    if (firstTagRunIndex >= 0 && firstTagRunIndex < runs.size()) {
                        formattingSourceRun = runs.get(firstTagRunIndex);
                    }

                    // Create a new run for the replacement with the formatting of the first tag run
                    XWPFRun replacementRun = paragraph.createRun();
                    if (formattingSourceRun != null) {
                        copyRunFormatting(formattingSourceRun, replacementRun);
                    }
                    replacementRun.setText(newReplacement);
                    System.err.println("Replacement Text" + replacementRun.text());
                    // Remove the runs that contained the tag (iterate in reverse)
                    for (int j = runs.size() - 1; j >= 0; j--) {
                        if (j >= firstTagRunIndex && j <= lastTagRunIndex) {
                            paragraph.removeRun(j);
                        }
                    }

                    // Insert the replacement run at the position of the first tag run
                    List<XWPFRun> currentRuns = paragraph.getRuns();
                    if (firstTagRunIndex <= currentRuns.size()) {
                        paragraph.addRun( replacementRun);
                    } else {
                        paragraph.addRun(replacementRun); // Add at the end if index is out of bounds (shouldn't happen)
                    }
                }
            }
        }

    }

    //Process standard tags
    public static Object getNestedValue(Map<String, Object> data, String path) {
        List<String> keys = Arrays.asList(path.split("\\."));
        return getNestedValueRecursive(data, keys, 0, path);
    }

    private static Object getNestedValueRecursive(Map<String, Object> currentLevel, List<String> keys, int index, String path) {
        if (currentLevel == null || index >= keys.size()) {
            System.err.println("Invalid path or null map at index: " + index);
            return null; // Return null if the path is invalid
        }

        String currentKey = keys.get(index);
        Object value = currentLevel.get(currentKey);

        if (value == null) {
            System.err.println("Key not found: " + currentKey);
            return "{{"+path+"}}"; // Key not found
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
        }else  {
            System.err.println("Intermediate value is not a map or list for key: " +
                    currentKey);
            return null; // Path not found or intermediate level is not a map or list
        }
    }
}


