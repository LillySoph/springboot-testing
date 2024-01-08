package lilly.testing.springboottesting;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

public class DocCreator {

    private File file;
    private String outputName;
    private XWPFDocument document;
    private Map<String, String> newContent;

    public void loadDocument(String template) throws IOException {
        // check if template file exists
        this.file = new File(template);
        if (!file.exists()) {
            throw new FileNotFoundException("Die Vorlage '" + template + "' konnte nicht gefunden werden.");
        }
        this.outputName = template.replace(".docx", "_filled.docx");
        createTestData();
        // load template file into XWPF doc
        try(FileInputStream fis = new FileInputStream(this.file)) {
            this.document = new XWPFDocument(fis);
            //parseDocument();
            updateDocument();
        }
    }

    private void createTestData() {
        this.newContent = new HashMap<String, String>();
        this.newContent.put("datum", "01.12.3005");
        this.newContent.put("betreff", "Das ist der Betreff");
        this.newContent.put("tgbNr", "12345");
        this.newContent.put("hier1", "Das wird ziemlich lecker.");
        this.newContent.put("hier2", "LeckerLeckerLecker");
        this.newContent.put("istGeschenk", "Wird gebacken für die Abschlussfeier von Cola");
        this.newContent.put("rezept", "LeckerLecker YumYum");
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

    private void replaceVariable(XWPFRun xwpfRun) {
        // check if variable exists and replace with content
        String runText = xwpfRun.text();
        System.out.println("\n--> Replace variable in Run '" + runText + "'");
        for (String key : newContent.keySet()) {
            String variable = "$" + key;
            System.out.println("Looking for variable " + variable);
            if (runText.contains(variable)) {
                System.out.println("Replace '" + key + "' with '" + newContent.get(key) + "'");
                runText = runText.replace(variable, newContent.get(key));
            }
        }
        xwpfRun.setText(runText, 0);
        System.out.println("--> Run after replacement: \n" + xwpfRun.text());
        //return xwpfRun;
    }

    private void replaceVariablesInParagraphs(List<XWPFParagraph> xwpfParagraphs) {
        for (XWPFParagraph xwpfParagraph : xwpfParagraphs) {
            List<XWPFRun> xwpfRuns = xwpfParagraph.getRuns();
            for(XWPFRun xwpfRun : xwpfRuns) {
                replaceVariable(xwpfRun);
            }
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

    public void replaceVariables(List<IBodyElement> bodyElements) {
        for (IBodyElement element : bodyElements) {
            if (element instanceof XWPFParagraph) {
                replaceParagraph((XWPFParagraph) element);
                //for (XWPFRun run : ((XWPFParagraph) element).getRuns()) {
                    //replaceRun(run);
                //}
            }
            else if (element instanceof XWPFTable) {
                for (XWPFTableRow row : ((XWPFTable) element).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        //replaceCell(cell);
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                                replaceParagraph(paragraph);
                        }
                    }
                }
            }
        }
    }

    private void replaceParagraph(XWPFParagraph paragraph) {
        if (paragraph.getCTP().getBookmarkStartList().isEmpty()) {
            return;
        }
        //System.out.println("[" + paragraph.getText() + "]\n" + paragraph.getCTP());
        char asciiAsChar = (char) 8194;
        String content = "";
        boolean onlyVar = true;
        for (XWPFRun run : paragraph.getRuns()) {
            if (run != null) {
                if(!run.text().contains(""+ asciiAsChar) && !run.text().isEmpty())
                    onlyVar = false;
            }
        }
        if (onlyVar)
            System.out.println("-> Paragraph [" + paragraph.getText() + "] only contains variable");
        else
            System.out.println("-> Paragraph [" + paragraph.getText() + "] contains text");
        for (CTBookmark bookmark : paragraph.getCTP().getBookmarkStartList()) {
            System.out.println("--> BOOKMARK: [" + bookmark.getName() + "]");
            //System.out.println(bookmark.getName());
        }
        // return when empty or no variable found
    }

    private void replaceRun(XWPFRun run) {

    }

    private void replaceCell(XWPFTableCell cell) {
        for (XWPFParagraph paragraph : cell.getParagraphs()) {
            for (XWPFRun run : paragraph.getRuns()) {
                replaceRun(run);
            }
        }
    }

    /**
     * Replace variables in template with new content
     */
    public void updateDocument() throws IOException {
        // replace variables in header
        for (XWPFHeader header : this.document.getHeaderList()) {
            replaceVariables(header.getBodyElements());
        }
        // replace variables in main document
        replaceVariables(this.document.getBodyElements());
        // replace variables in footer
        for (XWPFFooter footer : this.document.getFooterList()) {
            replaceVariables(footer.getBodyElements());
        }
        File output = new File(outputName);
        if (output.exists()) {
            output.delete();
        }
        try (FileOutputStream fos = new FileOutputStream(outputName)) {
                this.document.write(fos);
            }
    }

}
