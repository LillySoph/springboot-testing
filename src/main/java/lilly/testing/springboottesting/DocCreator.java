package lilly.testing.springboottesting;

import org.apache.poi.xwpf.usermodel.*;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;

import java.io.*;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DocCreator {

    private File file;
    private String outputName;
    private XWPFDocument document;
    private static String PATH = "d:\\Projekte\\SpringBootTesting\\generated\\";
    private JSONObject variableReplacement;
    private JSONArray tableContent;

    public void fillTemplate(String template) throws IOException {
        // check if template file exists
        this.file = new File(template);
        if (!file.exists()) {
            throw new FileNotFoundException("Die Vorlage '" + template + "' konnte nicht gefunden werden.");
        }
        this.outputName = template.replace(".docx", "_filled.docx");
        // read data from JSON
        this.variableReplacement = parseJSONObject(PATH + "data.json");
        this.tableContent = parseJSONArray(PATH + "formatted_data.json");
        // load template file into XWPF doc
        try(FileInputStream fis = new FileInputStream(this.file)) {
            this.document = new XWPFDocument(fis);
            updateDocument();
        }
    }

    /**
     * Replace variables and fill table in template
     */
    public void updateDocument() throws IOException {
        // replace variables in header
        for (XWPFHeader header : this.document.getHeaderList()) {
            insertReplacements(header.getBodyElements());
        }
        // replace variables in main document
        insertReplacements(this.document.getBodyElements());
        // replace variables in footer
        for (XWPFFooter footer : this.document.getFooterList()) {
            insertReplacements(footer.getBodyElements());
        }
        File output = new File(outputName);
        if (output.exists()) {
            output.delete();
        }
        try (FileOutputStream fos = new FileOutputStream(outputName)) {
            this.document.write(fos);
        }
    }

    public void insertReplacements(List<IBodyElement> bodyElements) {
        for (IBodyElement element : bodyElements) {
            if (element instanceof XWPFParagraph) {
                replaceVariableInParagraph((XWPFParagraph) element);
            } else if (element instanceof XWPFTable) {
                insertTableContent((XWPFTable) element);
            }
        }
    }

    private void replaceVariableInParagraph(XWPFParagraph paragraph) {
        // check if variable exists in paragraph
        if (!paragraph.getText().contains("$")) {
            return;
        }
        // extract style from first run in paragraph
        boolean isBold = paragraph.getRuns().get(0).isBold();
        boolean isItalic = paragraph.getRuns().get(0).isItalic();
        String fontColor = paragraph.getRuns().get(0).getColor();
        int fontSize = 11;
        // find variable in paragraph via RegEx matching
        String paragraphText = paragraph.getParagraphText();
        Pattern pattern = Pattern.compile("\\$\\w+");
        Matcher matcher = pattern.matcher(paragraphText);
        // replace variable with data
        while (matcher.find()) {
            String foundMatch = matcher.group();
            String replacement = getData(foundMatch.substring(1));
            if (replacement != null)
                paragraphText = paragraphText.replace(foundMatch, replacement);
        }
        clearParagraph(paragraph);
        // insert new paragraph content with formatting
        String[] paragraphTextSplitted = paragraphText.split(" ");
        for (int i = 0; i < paragraphTextSplitted.length; i++) {
            XWPFRun newRun = paragraph.insertNewRun(i);
            newRun.setBold(isBold);
            newRun.setFontSize(fontSize);
            newRun.setColor(fontColor);
            newRun.setItalic(isItalic);
            if (i + 1 < paragraphTextSplitted.length)
                newRun.setText(paragraphTextSplitted[i] + " ");
            else
                newRun.setText(paragraphTextSplitted[i]);
        }
    }

    private String getData(String key) {
        if (variableReplacement.get(key) == null)
            return null;
        return variableReplacement.get(key).toString();
    }

    private JSONObject parseJSONObject(String filename) {
        JSONParser jsonParser = new JSONParser();
        try (FileReader reader = new FileReader(filename)) {
            Object object = jsonParser.parse(reader);
            JSONObject jsonObject = (JSONObject) object;
            return jsonObject;
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
        return null;
    }

    private JSONArray parseJSONArray(String filename) {
        JSONParser jsonParser = new JSONParser();
        try (FileReader reader = new FileReader(filename)) {
            Object object = jsonParser.parse(reader);
            JSONArray jsonArray = (JSONArray) object;
            return jsonArray;
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }
        return null;
    }

    private void insertTableContent(XWPFTable table) {
        String key = "freitext";
        for (Object item : tableContent) {
            JSONObject jsonRow = (JSONObject) item;
            XWPFTableRow row = table.createRow();
            row.getCell(0).setText(jsonRow.get("zeit").toString());
            insertFormattedTableContent(row.getCell(1), jsonRow.get(key));
        }
    }

    private void insertFormattedTableContent(XWPFTableCell cell, Object formattedText) {
        // Format as described by the Quill text editor: https://quilljs.com/
        // -> "text": text, "attributes": [ array of formatting attributes of text ]
        try {
            JSONArray jsonArray = (JSONArray) formattedText;
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject jsonObject = (JSONObject) jsonArray.get(i);
                XWPFRun run = paragraph.insertNewRun(i);
                createFormattedRun(run, jsonObject);
                System.out.println(jsonArray.get(0));
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

    }

    private XWPFRun createFormattedRun(XWPFRun run, JSONObject data) {
        run.setText(data.get("text").toString());
        JSONObject attributes = (JSONObject) data.get("attributes");
        if (attributes != null) {
            // bold
            if (attributes.get("bold") != null)
                run.setBold((boolean) attributes.get("bold"));
            // italic
            if (attributes.get("italic") != null)
                run.setItalic((boolean) attributes.get("italic"));
            // underline
            if (attributes.get("underline") != null && ((boolean) attributes.get("underline")))
                run.setUnderline(UnderlinePatterns.SINGLE);
        }
        return run;
    }

    private XWPFParagraph clearParagraph(XWPFParagraph paragraph) {
        int size = paragraph.getRuns().size();
        for (int i = 0; i < size; i++) {
            paragraph.removeRun(0);
        }
        return paragraph;
    }

}
