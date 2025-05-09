package ddbmerge;

import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class ProcessHeadersFooters {

    public static void main(String[] args) {
        String inputFilePath = "C:\\Users\\hayde\\OneDrive\\Desktop\\Projects\\MergeDocs\\TestMerge.docx";
        String outputFilePath = "output_merged.docx";
        String tagToReplace = "{{TestMerge}}";
        String replacementText = "Hello World";
        String headerReplacementText = "Header Text";
        String footerReplacementText = "Footer Text";
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
        try {
            XWPFDocument document = readDocxFile(inputFilePath);

            if (document != null) {
                // Replace tag in main document paragraphs
                for (XWPFParagraph paragraph : document.getParagraphs()) {
                    mergeTagInParagraph(paragraph, tagToReplace, replacementText);
                }

                // Replace tag in headers
                for (XWPFHeader header : document.getHeaderList()) {
                    for (XWPFParagraph paragraph : header.getParagraphs()) {
                        mergeTagInParagraph(paragraph, tagToReplace, headerReplacementText);
                    }
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

    // First pass: Find the tag and the indices of the runs it spans
    for (int i = 0; i < runs.size(); i++) {
        XWPFRun run = runs.get(i);
        String runText = run.getText(0) == null ? "" : run.getText(0);
        int startIndexInParagraph = paragraphText.length();
        paragraphText.append(runText);
        int endIndexInParagraph = paragraphText.length();

        if (tagStartIndex == -1 && paragraphText.toString().contains(tag)) {
            tagStartIndex = paragraphText.indexOf(tag);
            firstTagRunIndex = i;
        }

        if (tagStartIndex != -1 && endIndexInParagraph >= tagStartIndex + tag.length() && lastTagRunIndex == -1) {
            lastTagRunIndex = i;
            tagEndIndex = tagStartIndex + tag.length();
        }
    }

    if (tagStartIndex != -1) {
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
        replacementRun.setText(replacement);

        // Remove the runs that contained the tag (iterate in reverse)
        for (int i = runs.size() - 1; i >= 0; i--) {
            if (i >= firstTagRunIndex && i <= lastTagRunIndex) {
                paragraph.removeRun(i);
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


