package lilly.testing.springboottesting;

import org.apache.poi.xwpf.usermodel.*;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.JSONParser;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTMarkupRange;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class DocCreator {

    private File file;
    private String outputName;
    private XWPFDocument document;
    private Map<String, String> headerContent;
    private JSONArray tableContent;

    public void loadDocument(String template) throws IOException {
        // check if template file exists
        this.file = new File(template);
        if (!file.exists()) {
            throw new FileNotFoundException("Die Vorlage '" + template + "' konnte nicht gefunden werden.");
        }
        this.outputName = template.replace(".docx", "_filled.docx");
        parseTableContent();
        createTestData();
        // load template file into XWPF doc
        try(FileInputStream fis = new FileInputStream(this.file)) {
            this.document = new XWPFDocument(fis);
            //parseDocument();
            updateDocument();
        }
    }

    private void createTestData() {
        this.headerContent = new HashMap<>();
        this.headerContent.put("Datum", "01.12.3005");
        this.headerContent.put("Betreff", "Das ist der Betreff");
        this.headerContent.put("TgbNr", "12345");
        this.headerContent.put("hier1", "Das wird ziemlich lecker.");
        this.headerContent.put("hier2", "LeckerLeckerLecker");
        this.headerContent.put("Bezug", "Wird gebacken für die Abschlussfeier von Cola");
        this.headerContent.put("verdeckteErhebung", "LeckerLecker YumYum");
        this.headerContent.put("istpraeventiv", "ist präventiv");
        this.headerContent.put("Ziel", "Lilly");
    }

    /**
     * Read and print document
     */
    public void parseDocument() {
        System.out.println("PARSING HEADER");
        for (XWPFHeader header : this.document.getHeaderList()) {
            parseBodyElements(header.getBodyElements());
        }
        System.out.println("PARSING CONTENT");
        parseBodyElements(this.document.getBodyElements());
        System.out.println("PARSING FOOTER");
        for (XWPFFooter footer : this.document.getFooterList()) {
            parseBodyElements(footer.getBodyElements());
        }
    }

    /**
     * Replace variables in template with new content
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

    public void parseBodyElements (List<IBodyElement> bodyElements) {
        for (IBodyElement element : bodyElements) {
            if (element instanceof XWPFParagraph) {
                for (XWPFRun run : ((XWPFParagraph) element).getRuns()) {
                    System.out.println("run:" + run.getText(0));
                }
            }
            else if (element instanceof XWPFTable) {
                for (XWPFTableRow row : ((XWPFTable) element).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        System.out.println("cell:" + cell.getText());
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            for (XWPFRun run : paragraph.getRuns()) {
                                System.out.println("- run in cell:" + run.getText(0));
                            }
                        }
                    }
                }
            }
        }
    }

    public void insertReplacements(List<IBodyElement> bodyElements) {
        for (IBodyElement element : bodyElements) {
            if (element instanceof XWPFParagraph) {
                replaceVariableInParagraph((XWPFParagraph) element);
                //replaceVariables((XWPFParagraph) element);
            }
            else if (element instanceof XWPFTable) {
                insertTableContent((XWPFTable) element);
                /*for (XWPFTableRow row : ((XWPFTable) element).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            replaceVariables(paragraph);
                        }
                    }
                }*/
            }
        }
    }

    private void replaceVariableInParagraph(XWPFParagraph paragraph) {
        // check for variable
        if (!paragraph.getText().contains("$")) {
            //System.out.println("No variable found: [" + paragraph.getText() + "]");
            return;
        }
        XWPFRun firstRun = paragraph.getRuns().get(0);
        boolean isBold = firstRun.isBold();
        boolean isItalic = firstRun.isItalic();
        //System.out.println("text scale: " + firstRun.getTextScale() + " | size: " + firstRun.getFontSize());
        int fontSize = 12;
        String fontColor = firstRun.getColor();
        List<XWPFRun> runList = paragraph.getRuns();
        String paragraphText = paragraph.getParagraphText();
        Pattern pattern = Pattern.compile("\\$\\w+");
        Matcher matcher = pattern.matcher(paragraphText);
        String variable, findings;
        System.out.println("paragraph before replacing variables:\n[" + paragraphText + "]");
        while (matcher.find()) {
            //System.out.println("Matching: [" + matcher.group() + "]");
            findings = matcher.group();
            variable = findings.substring(1);
            if (headerContent.get(variable) != null)
                paragraphText = paragraphText.replace(findings, headerContent.get(variable));
        }
        System.out.println("paragraph after replacing variables:\n[" + paragraphText + "]");
        clearParagraph(paragraph);
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

    private void replaceVariables(XWPFParagraph paragraph) {
        // check for variable
        if (!paragraph.getText().contains("$")) {
            //System.out.println("No variable found: [" + paragraph.getText() + "]");
            return;
        }
        String variableTrimmed = "";
        System.out.println("---");
        for (XWPFRun run : paragraph.getRuns()) {
            System.out.println("Run: [" + run.getText(0) + "]");
            if (run != null && run.getText(0) != null && run.getText(0).matches("\\$\\w+")) {
                // extract variable from run without the '$'
                variableTrimmed = run.getText(0).substring(1);
                System.out.println("\nvariable [" + variableTrimmed + "]");
                // check for replacement
                String bookmarkReplacement = headerContent.get(variableTrimmed);
                if (bookmarkReplacement == null) {
                    return;
                }
                System.out.println("Run before replacement: \n[" + run.getText(0) + "]");
                run.setText(bookmarkReplacement, 0);
                System.out.println("Run after replacement: \n[" + run.getText(0) + "]");
            }
        }
        System.out.println("Paragraph after replacement: \n[" + paragraph.getText() + "]");
    }

    private void parseTableContent() {
        System.out.println("--- Trying to parse JSON file");
        JSONParser jsonParser = new JSONParser();
        try (FileReader reader = new FileReader("d:\\Projekte\\SpringBootTesting\\generated\\formatted_data.json")) {
            Object object = jsonParser.parse(reader);
            this.tableContent = (JSONArray) object;
            for (Object item : tableContent) {
                JSONObject jsonRow = (JSONObject) item;
                System.out.println("-\nZEIT: " + jsonRow.get("zeit"));
                JSONArray freitext = (JSONArray) jsonRow.get("freitext");
                for (Object piece : freitext) {
                    JSONObject bestandteil = (JSONObject) piece;
                    System.out.println("BESTANDTEIL: " + bestandteil);
                    if (bestandteil.get("attributes") != null) {
                        System.out.println("Es hat attribute");
                    }
                }
                //System.out.println("FREITEXT: " + row.get("freitext"));
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

        //todo
    }

    private void insertTableContent(XWPFTable table) {

        for (Object item : tableContent) {
            JSONObject jsonRow = (JSONObject) item;
            XWPFTableRow row = table.createRow();
            row.getCell(0).setText(jsonRow.get("zeit").toString());
            insertFormattedText(row.getCell(1), jsonRow.get("freitext"));
            //row.getCell(1).setText(jsonRow.get("freitext").toString());

            //System.out.println("FREITEXT: " + row.get("freitext"));
        }

    }

    private void insertFormattedText(XWPFTableCell cell, Object formattedText) {
        // Format as described by the Quill text editor: https://quilljs.com/
        // -> "text": text, "attributes": [ array of format attributes of text ]
        try {
            JSONArray jsonArray = (JSONArray) formattedText;
            XWPFParagraph paragraph = cell.getParagraphs().get(0);
            for (int i = 0; i < jsonArray.size(); i++) {
                JSONObject jsonObject = (JSONObject) jsonArray.get(i);
                XWPFRun run = paragraph.insertNewRun(i);
                run.setText(jsonObject.get("text").toString());
                if (jsonObject.get("attributes") != null) {
                    JSONObject attributes = (JSONObject) jsonObject.get("attributes");
                    // bold
                    if (attributes.get("bold") != null)
                        run.setBold((boolean) attributes.get("bold"));
                    // italic
                    if (attributes.get("italic") != null)
                        run.setItalic((boolean) attributes.get("italic"));
                    // underline
                    //if (attributes.get("underline") != null)
                    //    run.setUnderline((boolean) attributes.get("underline"));
                }
                System.out.println(jsonArray.get(0));
            }
        } catch (Exception e) {
            System.out.println(e.getMessage());
        }

    }

    private XWPFParagraph clearParagraph(XWPFParagraph paragraph) {
        int size = paragraph.getRuns().size();
        for (int i = 0; i < size; i++) {
            paragraph.removeRun(0);
        }
        return paragraph;
    }

    private void replaceBookmarks(XWPFParagraph paragraph) {
        // check for bookmark
        List<CTBookmark> bookmarkStartList = paragraph.getCTP().getBookmarkStartList();
        if (bookmarkStartList.isEmpty()) {
            return;
        }
        List<CTMarkupRange> markupRangeList = paragraph.getCTP().getBookmarkEndList();
        for (CTBookmark bookmark : bookmarkStartList) {
            System.out.println("---\nbookmark: [" + bookmark.getName() + "]");
            if (bookmark.getName().equals("Betreff"))
                System.out.println(bookmark);
        }
        // extract paragraph text
        String paragraphText = paragraph.getParagraphText();
        System.out.println("paragraph text: [" + paragraphText + "]");
        // replace bookmark with new content
        char bookmarkAsChar = (char) 8194;
        char[] bookmarkRepresentation = {bookmarkAsChar, bookmarkAsChar, bookmarkAsChar, bookmarkAsChar, bookmarkAsChar};
        // check for replacement
        String bookmarkReplacement = headerContent.get(bookmarkStartList.get(0).getName());
        if (bookmarkReplacement == null) {
            return;
        }
        String replacedParagraphText = paragraphText.replace(new String(bookmarkRepresentation), bookmarkReplacement);
        // in case the bookmark is invisible
        if (replacedParagraphText.isEmpty()) {
            replacedParagraphText = bookmarkReplacement;
        }
        System.out.println("replaced paragraph text: [" + replacedParagraphText + "]");
        // remove old runs of paragraph
        int numberOfRuns = paragraph.getRuns().size();
        for (int i = 0; i < numberOfRuns; i++) {
            System.out.println("[" + paragraph.getRuns().get(i) + "] -> is Bold: " + paragraph.getRuns().get(i).isBold());
        }
        for (int i = 0; i < numberOfRuns; i++) {
            paragraph.removeRun(0);
        }
        System.out.println("after removing runs: [" + paragraph.getText() + "]");
        // add run with new content
        String[] replacementSplitOnCarriageReturn = replacedParagraphText.split(" ");
        for (int i = 0; i < replacementSplitOnCarriageReturn.length; i++) {
            XWPFRun newRun = paragraph.insertNewRun(i);
            if (i + 1 < replacementSplitOnCarriageReturn.length)
                newRun.setText(replacementSplitOnCarriageReturn[i] + " ");
            else
                newRun.setText(replacementSplitOnCarriageReturn[i]);
        }

        System.out.println("after adding new runs: [" + paragraph.getText() + "]");

        /*
        System.out.println("[" + paragraph.getText() + "]\n-> contains bookmark: " + bookmarkStartList.get(0).getName());
        List<XWPFRun> runList = paragraph.getRuns();
        int startIndex = -1;
        for (int i = 0; i < runList.size(); i++) {
            if (runList.get(0).text().contains("" + ((char) 8194))) {
                if (startIndex < 0)
                    startIndex = i;
                else
                    paragraph.removeRun(i);
            }
        }
        System.out.println("Paragraph after changes: \n[" + paragraph.getText() + "]");

        List<String> bookmarkList = new ArrayList<String>();
        for (CTBookmark bookmark : bookmarkStartList) {
            //System.out.println("[" + paragraph.getText() + "]\nParagraph contains bookmark: " + bookmark.getName());
            bookmarkList.add(bookmark.getName());
        }
        char bookmarkAsChar = (char) 8194; // bookmarks are represented by '?'
        int startposition
        for (int i = 0; i < paragraph.getRuns().size(); i++) {
            if (paragraph.getRuns().get(i).text().contains("" + bookmarkAsChar)) {
                paragraph.removeRun(i);
            }
        }
        boolean containsOnlyBookmark = true;
        // check for text in paragraph additionally to bookmark
        char[] bookmarkRepresentation = {bookmarkAsChar, bookmarkAsChar, bookmarkAsChar, bookmarkAsChar, bookmarkAsChar};
        String prefix = "";
        String bookmark = null;
        String suffix = "";
        for (XWPFRun run : paragraph.getRuns()) {
            if (run != null && (!run.text().contains("" + bookmarkAsChar) && !run.text().isEmpty())) {
                // paragraph contains text
                containsOnlyBookmark = false;
            }
        }
        if (!containsOnlyBookmark) {
            System.out.println("\n-> Paragraph [" + paragraph.getText() + "] ALSO contains text");
            String[] paragraphPieces = paragraph.getText().split(" ");
            for (String piece : paragraphPieces) {
                if (piece.contains("" + bookmarkAsChar)) {
                    bookmark = bookmarkStartList.get(0).getName();
                } else if (!piece.contains("" + bookmarkAsChar) && bookmark == null) {
                    prefix += piece + " ";
                } else if (bookmark != null){
                    suffix += piece + " ";
                }
            }
            System.out.print("result: [" + prefix + "], [" + bookmark + "], [" + suffix + "]\n");
            //replaceBookmarkWithText(paragraph);
        } else {
            bookmark = bookmarkStartList.get(0).getName();
        }
        replaceBookmark(paragraph, prefix, suffix, bookmark);
        */
    }

    private void replaceBookmark(XWPFParagraph paragraph, String prefix, String suffix, String bookmark) {
        for (int i = 0; i < paragraph.getRuns().size(); i++) {
            //System.out.println("replaceBookmark / run: " + paragraph.getRuns().get(i));
            paragraph.removeRun(i);
        }
        String replacement = this.headerContent.get(bookmark);
        if (replacement == null) {
            System.out.println("[" + bookmark + "] could not be replaced");
            return;
        }
        String newRun = "" + prefix + replacement + suffix;
        XWPFRun prefixRun = paragraph.insertNewRun(0);
        prefixRun.setText(prefix);
        XWPFRun replacementRun = paragraph.insertNewRun(1);
        replacementRun.setText(replacement);
        XWPFRun suffixRun = paragraph.insertNewRun(2);
        suffixRun.setText(suffix);
        System.out.println("new paragraph: " + paragraph.getText());
        //paragraph.addRun(new XWPFRun());

        // replace bookmark
    }



}
