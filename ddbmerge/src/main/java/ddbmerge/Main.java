package com.example;
 
import org.docx4j.XmlUtils;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.relationships.Relationship;
import org.docx4j.wml.*;
 
import com.fasterxml.jackson.databind.ObjectMapper;
 
import com.fasterxml.jackson.databind.util.JSONPObject;
 
//import jakarta.xml.bind.*;
 
import javax.xml.bind.*;
import javax.xml.namespace.QName;
 
import java.io.File;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.lang.reflect.Field;
import java.util.HashMap;
import java.util.List;
import java.util.ArrayList;
import java.util.Arrays;
 
public class Main {
    public static void oldJava(String[] args) throws Exception {
        // generate and setup test data
        String templatePath = "C:\\Users\\Hayden Work\\OneDrive - Dudley DeBosier, APLC\\DDBProjects\\Merge Docs\\src\\Matter Templates\\TestDocgen.docx";
        Map<String, String> objectMap = new HashMap<>();
        objectMap.put("Matter", "Matter Templates");
        objectMap.put("Intake", "Intake Templates");
        objectMap.put("Request", "Request Templates");
        objectMap.put("Accounting", "Accounting Templates");
        objectMap.put("Testing", "Testing Folder");
 
        String objectPath = objectMap.get("Matter");
        List<String> templates = new ArrayList<>();
        templates.add("TestDocgen.docx");
 
        // process JSON
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
            ObjectMapper objectMapper = new ObjectMapper();
            objectMapper.configure(com.fasterxml.jackson.core.JsonParser.Feature.ALLOW_SINGLE_QUOTES, true);
 
            Map<String, Object> processedJSON = objectMapper.readValue(json, Map.class);
 
            // Extracting "data"
            Map<String, Object> dataMap = (Map<String, Object>) processedJSON.get("data");
            try {
                process_doc(templatePath, dataMap, "docx");
 
            } catch (Exception e) {
                // TODO: handle exception
            }
        } catch (Exception e) {
            System.err.println("Error processing JSON: " + e.getMessage());
        }
    }
 
    public static void process_doc(String templatePath, Map<String, Object> dataMap, String output_format) {
        try {
            // Load the document
            File templateDocument = new File(templatePath);
            WordprocessingMLPackage newDocument = WordprocessingMLPackage.createPackage();
 
            WordprocessingMLPackage wordPackage = WordprocessingMLPackage.load(templateDocument);
 
            MainDocumentPart newMainPart = newDocument.getMainDocumentPart();
            MainDocumentPart mainDocumentPart = wordPackage.getMainDocumentPart();
 
            try {
                Boolean refreshXML = true;
 
                // Process the body
                try {
                    List<Object> documentBodies = mainDocumentPart.getJAXBNodesViaXPath("//w:body", refreshXML);
                    refreshXML = false;
                    for (Object documentBody : documentBodies) {
                        Body body = (Body) documentBody;
                        for (Object element : body.getContent()) {
                            if (element instanceof P) {
                                newMainPart.addObject(process_paragraph(element, dataMap)); // Add the updated P object
                            } else if (element instanceof JAXBElement<?>) {
                                newMainPart.addObject(process_table(element, dataMap));
                            }
                        }
                    }
                } catch (Exception e) {
                    System.err.println("Error Processing Body: " + e);
                }
 
            } catch (Exception e) {
                System.err.println("Error processing document structure: " + e);
            }
 
            // Create the document
            try {
                // Set JAXB context factory
                System.setProperty("javax.xml.bind.context.factory", "org.eclipse.persistence.jaxb.JAXBContextFactory");
 
                File outputDoc = new File("outputDoc.docx");
                newDocument.save(outputDoc);
                System.out.println("Document saved");
            } catch (Exception e) {
                System.err.println("Error Generating document: " + e);
            }
 
        } catch (Exception e) {
            System.err.println("Error trying to instantiate document: " + e);
        }
    }
 
    public static Tbl process_table(Object table, Map<String, Object> dataMap) {
        Tbl tbl = new Tbl();
        try {
            Object value = ((JAXBElement<?>) table).getValue();
            tbl = (Tbl) value;
            System.out.println("Processing Table");
 
            // Create a copy of the rows to avoid ConcurrentModificationException
            List<Object> rowsCopy = new ArrayList<>(tbl.getContent());
            if (rowsCopy != null) {
                for (Object row : rowsCopy) {
                    try {
                        process_tableRow(row, tbl, dataMap);
                    } catch (Exception e) {
                        System.err.println("Error processing table row: " + e);
                    }
                }
            }
            System.out.println("Done Processing Table");
 
        } catch (Exception e) {
            System.err.println("Error processing table: " + e);
        }
        return tbl;
    }
 
    public static void process_tableRow(Object row, Tbl parentTable, Map<String, Object> dataMap) {
        try {
            Tr value = (Tr) row; // Cast the row to Tr (table row)
            List<Object> cells = value.getContent(); // Get the cells of the row
 
            if (cells != null) {
                // Check if the first cell contains a {startRow} tag
                String firstCellContent = getFirstCellContent(row);
                if (firstCellContent.contains("startRow")) {
                    String regex = "\\{startRow\\.([\\w\\.]+)\\}";
                    Pattern pattern = Pattern.compile(regex);
                    Matcher matcher = pattern.matcher(firstCellContent);
 
                    List<Tr> replicatedRows = new ArrayList<>();
                    boolean removeOriginalRow = false;
 
                    while (matcher.find()) {
                        String matchedTag = matcher.group(1); // Extract the key after "startRow."
                        System.out.println("Key: " + matchedTag);
 
                        // Process the rows to replicate based on the matched key
                        replicatedRows.addAll(processReplicateRows(value, dataMap, matchedTag));
                        removeOriginalRow = true; // Mark the original row for removal
                    }
 
                    // Modify the table content after iteration
                    if (removeOriginalRow) {
                        int rowIndex = parentTable.getContent().indexOf(row);
                        parentTable.getContent().addAll(rowIndex + 1, replicatedRows); // Add replicated rows
                        parentTable.getContent().remove(row); // Remove the original template row
                    }
 
                    // Iterate over the replicated rows to process regular dataMap tags
                    for (Tr replicatedRow : replicatedRows) {
                        List<Object> replicatedCells = replicatedRow.getContent();
                        for (Object cell : replicatedCells) {
                            try {
                                process_RowCell(cell, dataMap); // Process each cell to replace regular tags
                            } catch (Exception e) {
                                System.err.println("Error processing replicated table cell: " + e);
                            }
                        }
                    }
                } else {
                    // Process each cell in the row for regular dataMap replacements
                    for (Object cell : cells) {
                        try {
                            process_RowCell(cell, dataMap);
                        } catch (Exception e) {
                            System.err.println("Error processing table cell: " + e);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Error processing table row: " + e);
        }
    }
 
    public static void process_FirstCell(Object cell, Map<String, Object> dataMap) {
        try {
            Object value = ((JAXBElement<?>) cell).getValue();
            Tc tc = (Tc) value; // Table cell
            List<Object> cellContents = tc.getContent();
            List<Object> updatedContent = new ArrayList<>();
 
            // Step 1: Process each paragraph in the first cell
            if (cellContents != null) {
                for (Object cellContent : cellContents) {
                    try {
                        if (cellContent instanceof P) {
                            // Get the paragraph text
                            String paragraphText = getParagraphText(cellContent);
 
                            // Replace startRow tags with blank space
                            String replacedText = replace_ReplicateTags(paragraphText, dataMap);
 
                            // Replace other tags
                            replacedText = replace_tags(replacedText, dataMap);
 
                            // Create a new paragraph with the replaced text
                            P updatedParagraph = createParagraphWithText(replacedText);
                            updatedContent.add(updatedParagraph);
                        } else {
                            // Preserve non-paragraph content (e.g., tables, drawings)
                            updatedContent.add(cellContent);
                        }
                    } catch (Exception e) {
                        System.err.println("Error processing first cell content: " + e);
                    }
                }
            }
 
            // Step 2: Replace the cell content with the updated paragraphs
            tc.getContent().clear();
            tc.getContent().addAll(updatedContent);
 
        } catch (Exception e) {
            System.err.println("Error processing first cell: " + e);
        }
    }
 
    public static void process_RowCell(Object cell, Map<String, Object> dataMap) {
        try {
            Object value = ((JAXBElement<?>) cell).getValue();
            Tc tc = (Tc) value; // Table cell
            List<Object> cellContents = tc.getContent();
            List<Object> updatedContent = new ArrayList<>();
 
            // Step 1: Process each paragraph in the cell
            if (cellContents != null) {
                for (Object cellContent : cellContents) {
                    try {
                        if (cellContent instanceof P) {
                            // Process the paragraph and replace tags
                            P updatedParagraph = process_paragraph(cellContent, dataMap);
                            updatedContent.add(updatedParagraph);
                        } else {
                            // Preserve non-paragraph content (e.g., tables, drawings)
                            updatedContent.add(cellContent);
                        }
                    } catch (Exception e) {
                        System.err.println("Error processing cell content: " + e);
                    }
                }
            }
 
            // Step 2: Replace the cell content with the updated paragraphs
            tc.getContent().clear();
            tc.getContent().addAll(updatedContent);
 
        } catch (Exception e) {
            System.err.println("Error processing table cell: " + e);
        }
    }
 
    private static P createParagraphWithText(String text) {
        ObjectFactory factory = new ObjectFactory();
        P paragraph = factory.createP();
        R run = factory.createR();
        Text t = factory.createText();
        t.setValue(text);
        run.getContent().add(t);
        paragraph.getContent().add(run);
        return paragraph;
    }
 
    public static P process_paragraph(Object paragraph, Map<String, Object> dataMap) {
        try {
            P p = (P) paragraph;
            List<Object> runs = p.getContent();
            StringBuilder fullText = new StringBuilder();
 
            // Step 1: Concatenate all text from the runs
            List<R> originalRuns = new ArrayList<>();
            if (runs != null) {
                for (Object run : runs) {
                    if (run instanceof R) {
                        R r = (R) run;
                        originalRuns.add(r);
                        List<Object> texts = r.getContent();
                        for (Object text : texts) {
                            if (text instanceof JAXBElement<?>) {
                                Object textValue = ((JAXBElement<?>) text).getValue();
                                if (textValue instanceof Text) {
                                    fullText.append(((Text) textValue).getValue());
                                }
                            }
                        }
                    }
                }
            }
 
            // Step 2: Replace {startRow...} tags with an empty string
            String replacedText = fullText.toString().replaceAll("\\{startRow\\.[\\w\\.]+\\}", "");
 
            // Step 3: Replace other tags in the concatenated text
            replacedText = replace_tags(replacedText, dataMap);
 
            // Step 4: Clear the original runs and rebuild them with the replaced text
            p.getContent().clear();
            if (!originalRuns.isEmpty()) {
                String[] splitText = replacedText.split("(?<=\\G.{50})"); // Split into chunks of 50 characters
                for (String chunk : splitText) {
                    R newRun = XmlUtils.deepCopy(originalRuns.get(0)); // Clone the first run to preserve formatting
                    List<Object> newRunContent = newRun.getContent();
                    newRunContent.clear();
 
                    Text newText = new Text();
                    newText.setValue(chunk);
                    JAXBElement<Text> textElement = new JAXBElement<>(
                            new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "t"), Text.class,
                            newText);
                    newRunContent.add(textElement);
 
                    p.getContent().add(newRun);
                }
            } else {
                // Handle empty paragraphs by creating a new run
                R newRun = new R();
                Text newText = new Text();
                newText.setValue(replacedText);
                JAXBElement<Text> textElement = new JAXBElement<>(
                        new QName("http://schemas.openxmlformats.org/wordprocessingml/2006/main", "t"), Text.class,
                        newText);
                newRun.getContent().add(textElement);
                p.getContent().add(newRun);
            }
 
            return p; // Return the updated paragraph
        } catch (Exception e) {
            System.err.println("Error processing paragraph: " + e);
        }
        return null; // Return null if processing fails
    }
 
    public static List<Object> process_run(Object run, Map<String, Object> dataMap) {
        List<Object> processedContent = new ArrayList<>();
        try {
            if (run instanceof R) {
                R r = (R) run;
                List<Object> texts = r.getContent();
 
                if (texts != null) {
                    for (Object text : texts) {
                        try {
                            if (text instanceof JAXBElement<?>) {
                                Object textValue = ((JAXBElement<?>) text).getValue();
                                if (textValue instanceof Text) {
                                    Text originalText = (Text) textValue;
                                    String replacedText = replace_tags(originalText.getValue(), dataMap);
 
                                    // Clone the original run to retain formatting
                                    R clonedRun = XmlUtils.deepCopy(r);
                                    List<Object> clonedRunContent = clonedRun.getContent();
 
                                    // Update the text in the cloned run
                                    for (Object clonedText : clonedRunContent) {
                                        if (clonedText instanceof JAXBElement<?>) {
                                            Object clonedTextValue = ((JAXBElement<?>) clonedText).getValue();
                                            if (clonedTextValue instanceof Text) {
                                                ((Text) clonedTextValue).setValue(replacedText);
                                            }
                                        }
                                    }
 
                                    processedContent.add(clonedRun);
                                }
                            }
                        } catch (Exception e) {
                            System.err.println("Error processing text: " + e);
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Error processing run: " + e);
        }
        return processedContent;
    }
 
    public static Object process_text(Object text) {
        try {
            Object t = ((JAXBElement<?>) text).getValue();
 
            if (t instanceof Text) {
                // Handle regular text
                return t; // Return the original Text object
            } else if (t instanceof Drawing) {
                // Handle images (Drawing objects)
                System.out.println("Found Drawing");
                Drawing drawing = (Drawing) t;
 
                // Clone the drawing to replicate it
                Drawing clonedDrawing = XmlUtils.deepCopy(drawing);
 
                // Return the cloned drawing
                return clonedDrawing;
            }
        } catch (Exception e) {
            System.err.println("Error processing text: " + e);
        }
        return null; // Return null if processing fails
    }
 
    public static String replace_tags(String text, Map<String, Object> dataMap) {
        try {
            // Regex to find tags in the format {key.subkey.subsubkey...}
            String regex = "\\{\\s*([\\w\\.]+)\\s*\\}";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(text);
 
            StringBuffer result = new StringBuffer();
            while (matcher.find()) {
                String fullKey = matcher.group(1).trim(); // Extract and trim the full key
                Object replacement = getNestedValue(dataMap, fullKey);
 
                if (replacement != null) {
                    matcher.appendReplacement(result, Matcher.quoteReplacement(replacement.toString()));
                    System.out.println("Replacing tag: " + fullKey + " with value: " + replacement);
                } else {
                    System.err.println("Key not found in dataMap: " + fullKey);
                    matcher.appendReplacement(result, Matcher.quoteReplacement(matcher.group(0))); // Keep the original
                                                                                                   // tag
                }
            }
            matcher.appendTail(result);
            return result.toString();
        } catch (Exception e) {
            System.err.println("Error replacing tags: " + e);
            return text; // Return original text in case of error
        }
    }
 
    public static Object getNestedValue(Map<String, Object> data, String path) {
        List<String> keys = Arrays.asList(path.split("\\."));
        return getNestedValueRecursive(data, keys, 0);
    }
 
    private static Object getNestedValueRecursive(Map<String, Object> currentLevel, List<String> keys, int index) {
        if (currentLevel == null || index >= keys.size()) {
            System.err.println("Invalid path or null map at index: " + index);
            return null; // Return null if the path is invalid
        }
 
        String currentKey = keys.get(index);
        Object value = currentLevel.get(currentKey);
 
        if (value == null) {
            System.err.println("Key not found: " + currentKey);
            return null; // Key not found
        }
 
        if (index == keys.size() - 1) {
            // If this is the last key, return the value
            return value;
        } else if (value instanceof Map) {
            // If the value is a map, continue recursively
            return getNestedValueRecursive((Map<String, Object>) value, keys, index + 1);
        } else if (value instanceof List) {
            // If the value is a list, get the first element and continue if it's a map
            List<?> list = (List<?>) value;
            if (!list.isEmpty() && list.get(0) instanceof Map) {
                return getNestedValueRecursive((Map<String, Object>) list.get(0), keys, index
                        + 1);
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
 
    public static String getFirstCellContent(Object row) {
        try {
            Tr value = (org.docx4j.wml.Tr) row;
            List<Object> cells = value.getContent();
            System.out.println("Processing Table Row");
 
            if (cells != null && !cells.isEmpty()) {
                Object firstCell = cells.get(0); // Get the first cell
                return getCellContentAsString(firstCell); // Extract and return its content
            }
        } catch (Exception e) {
            System.err.println("Error processing table row: " + e);
        }
        return null; // Return null if no content is found
    }
 
    private static String getCellContentAsString(Object cell) {
        try {
            Object value = ((JAXBElement<?>) cell).getValue();
            Tc tc = (Tc) value;
 
            StringBuilder cellText = new StringBuilder();
            List<Object> cellContents = tc.getContent();
 
            if (cellContents != null) {
                for (Object cellContent : cellContents) {
                    if (cellContent instanceof P) {
                        P paragraph = (P) cellContent;
                        List<Object> runs = paragraph.getContent();
                        if (runs != null) {
                            for (Object run : runs) {
                                if (run instanceof R) {
                                    R r = (R) run;
                                    List<Object> texts = r.getContent();
                                    if (texts != null) {
                                        for (Object text : texts) {
                                            if (text instanceof JAXBElement<?>) {
                                                Object textValue = ((JAXBElement<?>) text).getValue();
                                                if (textValue instanceof Text) {
                                                    cellText.append(((Text) textValue).getValue());
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return cellText.toString();
        } catch (Exception e) {
            System.err.println("Error extracting cell content: " + e);
        }
        return "";
    }
 
    private static List<Tr> processReplicateRows(Tr templateRow, Map<String, Object> dataMap,
            String repeatingObjectKey) {
        List<Tr> rows = new ArrayList<>();
 
        try {
            // Retrieve the list of objects from the dataMap
            List<Map<String, Object>> objectList = (List<Map<String, Object>>) dataMap.get(repeatingObjectKey);
            if (objectList == null || objectList.isEmpty()) {
                System.err.println("No data found for key: " + repeatingObjectKey);
                return rows; // Return empty list if no data is found
            }
 
            // Loop over the list of objects
            for (Map<String, Object> iterObject : objectList) {
                // Clone the template row
                Tr newRow = XmlUtils.deepCopy(templateRow);
 
                // Process each cell in the row
                List<Object> cells = newRow.getContent();
                for (Object cell : cells) {
                    try {
                        // Replace tags in the cell with data from iterObject
                        Object cellValue = ((JAXBElement<?>) cell).getValue();
                        Tc tc = (Tc) cellValue;
 
                        List<Object> cellContents = tc.getContent();
                        List<Object> updatedContent = new ArrayList<>();
 
                        if (cellContents != null) {
                            for (Object cellContent : cellContents) {
                                if (cellContent instanceof P) {
                                    // Clone the paragraph to retain formatting
                                    P originalParagraph = (P) cellContent;
                                    P clonedParagraph = XmlUtils.deepCopy(originalParagraph);
 
                                    // Step 1: Concatenate all text from the runs
                                    List<Object> runs = clonedParagraph.getContent();
                                    StringBuilder fullText = new StringBuilder();
                                    List<R> originalRuns = new ArrayList<>();
                                    for (Object run : runs) {
                                        if (run instanceof R) {
                                            R r = (R) run;
                                            originalRuns.add(r); // Save the original run for formatting
                                            List<Object> texts = r.getContent();
                                            for (Object text : texts) {
                                                if (text instanceof JAXBElement<?>) {
                                                    Object textValue = ((JAXBElement<?>) text).getValue();
                                                    if (textValue instanceof Text) {
                                                        fullText.append(((Text) textValue).getValue());
                                                    }
                                                }
                                            }
                                        }
                                    }
 
                                    // Step 2: Replace tags in the concatenated text
                                    String replacedText = replace_ReplicateTags(fullText.toString(), iterObject);
                                    replacedText = replace_tags(replacedText, iterObject);
 
                                    // Step 3: Rebuild the runs while preserving formatting
                                    clonedParagraph.getContent().clear();
                                    int currentIndex = 0;
                                    for (R originalRun : originalRuns) {
                                        List<Object> texts = originalRun.getContent();
                                        for (Object text : texts) {
                                            if (text instanceof JAXBElement<?>) {
                                                Object textValue = ((JAXBElement<?>) text).getValue();
                                                if (textValue instanceof Text) {
                                                    Text originalText = (Text) textValue;
                                                    String originalRunText = originalText.getValue();
 
                                                    // Determine the portion of the replaced text corresponding to this
                                                    // run
                                                    int endIndex = currentIndex + originalRunText.length();
                                                    if (endIndex > replacedText.length()) {
                                                        endIndex = replacedText.length();
                                                    }
                                                    String newRunText = replacedText.substring(currentIndex, endIndex);
                                                    currentIndex = endIndex;
 
                                                    // Clone the original run and update its text
                                                    R newRun = XmlUtils.deepCopy(originalRun);
                                                    List<Object> newRunContent = newRun.getContent();
                                                    newRunContent.clear();
 
                                                    Text newText = new Text();
                                                    newText.setValue(newRunText);
                                                    JAXBElement<Text> textElement = new JAXBElement<>(
                                                            new QName(
                                                                    "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
                                                                    "t"),
                                                            Text.class,
                                                            newText);
                                                    newRunContent.add(textElement);
 
                                                    // Add the updated run to the paragraph
                                                    clonedParagraph.getContent().add(newRun);
                                                }
                                            }
                                        }
                                    }
 
                                    updatedContent.add(clonedParagraph); // Add the updated paragraph
                                } else {
                                    // Preserve non-paragraph content (e.g., tables, drawings)
                                    updatedContent.add(cellContent);
                                }
                            }
                        }
 
                        // Replace the cell content with the updated paragraphs
                        tc.getContent().clear();
                        tc.getContent().addAll(updatedContent);
 
                    } catch (Exception e) {
                        System.err.println("Error processing cell: " + e);
                    }
                }
 
                // Add the updated row to the list of rows
                rows.add(newRow);
            }
 
        } catch (Exception e) {
            System.err.println("Error processing replicate rows: " + e);
        }
 
        return rows;
    }
 
    public static String replace_ReplicateTags(String text, Map<String, Object> objectMap) {
        try {
            // Regex to find tags in the format {key.subkey.subsubkey...}
            String regex = "\\{\\s*([\\w\\.]+)\\s*\\}";
            Pattern pattern = Pattern.compile(regex);
            Matcher matcher = pattern.matcher(text);
            StringBuffer result = new StringBuffer();
            while (matcher.find()) {
 
                String fullKey = matcher.group(1); // Extract the full key (e.g., startRow.OBJECT or Roles.Name)
                System.out.println("REGEX: " + fullKey);
 
                // Check if the tag starts with "startRow."
                if (fullKey.contains("startRow.")) {
                    // Replace {startRow.OBJECT} with blank text
                    matcher.appendReplacement(result, "");
                    System.out.println("Replacing Replicate tag: " + fullKey + " with blank space");
                } else {
                    // Handle other tags normally
                    Object replacement = getNestedValue(objectMap, fullKey);
                    if (replacement != null) {
                        matcher.appendReplacement(result, Matcher.quoteReplacement(replacement.toString()));
                        System.out.println("Replacing tag: " + fullKey + " with value: " + replacement);
                    } else {
                        // Keep the original tag if no replacement is found
                        matcher.appendReplacement(result, Matcher.quoteReplacement(matcher.group(0)));
                    }
                }
            }
            matcher.appendTail(result);
            return result.toString();
        } catch (Exception e) {
            System.err.println("Error replacing tags: " + e);
            return text; // Return original text in case of error
        }
    }
 
    private static String getParagraphText(Object paragraph) {
        StringBuilder fullText = new StringBuilder();
        try {
            if (paragraph instanceof P) {
                P p = (P) paragraph;
                List<Object> runs = p.getContent();
 
                if (runs != null) {
                    for (Object run : runs) {
                        if (run instanceof R) {
                            R r = (R) run;
                            List<Object> texts = r.getContent();
                            if (texts != null) {
                                for (Object text : texts) {
                                    if (text instanceof JAXBElement<?>) {
                                        Object textValue = ((JAXBElement<?>) text).getValue();
                                        if (textValue instanceof Text) {
                                            fullText.append(((Text) textValue).getValue());
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        } catch (Exception e) {
            System.err.println("Error extracting paragraph text: " + e);
        }
        return fullText.toString();
    }
 
}
 