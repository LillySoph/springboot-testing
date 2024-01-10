package lilly.testing.springboottesting;

import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBookmark;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Find all bookmarks in a given DOCX file and print it out to a TXT file to verify the template before using it
 */
public class TemplateVerifier {

    private File file;
    List<String> allBookmarks = new ArrayList<>();

    List<String> allVariables = new ArrayList<>();
    private XWPFDocument document;

    public void validateVariableTemplate(String fileName) throws IOException {
        // check if file exists
        this.file = new File(fileName);
        if (!file.exists()) {
            throw new FileNotFoundException("Die Datei '" + fileName + "' konnte nicht gefunden werden.");
        }
        // load file into XWPF doc, check for variables and output findings in txt
        try(FileInputStream fis = new FileInputStream(this.file)) {
            this.document = new XWPFDocument(fis);
            findVariables();
            File resultFile = new File("d:\\Projekte\\SpringBootTesting\\generated\\result_of_variable_check.txt");
            BufferedWriter writer = new BufferedWriter(new FileWriter(resultFile));
            writer.write("--- Die folgenden Variablen wurden in der Datei '" + fileName + "' gefunden ---");
            for (String variable : this.allVariables) {
                writer.write("\n" + variable);
            }
            writer.write("\n--- Sollte eine Variable fehlen, bitte die Vorlagendatei nochmal pruefen und ggf. neu einfuegen ---");
            writer.write("\n--- Bitte beachten: Pro Zeile darf nur eine Variable existieren ---");
            writer.close();
        }
    }

    private void findVariables() {
        // checking header
        for (XWPFHeader header : this.document.getHeaderList()) {
            checkBodyElementsForVariables(header.getBodyElements());
        }
        // checking document
        checkBodyElementsForVariables(this.document.getBodyElements());
        // checking footer
        for (XWPFFooter footer : this.document.getFooterList()) {
            checkBodyElementsForVariables(footer.getBodyElements());
        }
    }


    private void checkBodyElementsForVariables(List<IBodyElement> bodyElements) {
        for (IBodyElement element : bodyElements) {
            // check paragraphs
            if (element instanceof XWPFParagraph) {
                checkParagraphForVariables((XWPFParagraph) element);
            }
            // check tables
            else if (element instanceof XWPFTable) {
                for (XWPFTableRow row : ((XWPFTable) element).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            checkParagraphForVariables(paragraph);
                        }
                    }
                }
            }
        }
    }

    private void checkParagraphForVariables(XWPFParagraph paragraph) {
        for (XWPFRun run : paragraph.getRuns()) {
            System.out.println("RUN: [" + run.getText(0) + "]");
            if (run.text().matches("(\\$\\w+)")) {
                allVariables.add(run.text());
            }
        }
    }

    public void validateBookmarkTemplate(String fileName) throws IOException {
        // check if file exists
        this.file = new File(fileName);
        if (!file.exists()) {
            throw new FileNotFoundException("Die Datei '" + fileName + "' konnte nicht gefunden werden.");
        }
        // load file into XWPF doc, check for bookmarks and output findings in txt
        try(FileInputStream fis = new FileInputStream(this.file)) {
            this.document = new XWPFDocument(fis);
            findBookmarks();
            File resultFile = new File("d:\\Projekte\\SpringBootTesting\\generated\\result_of_bookmark_check.txt");
            BufferedWriter writer = new BufferedWriter(new FileWriter(resultFile));
            writer.write("--- Die folgenden Textmarken wurden in der Datei '" + fileName + "' gefunden ---");
            for (String bookmark : this.allBookmarks) {
                writer.write("\n" + bookmark);
            }
            writer.write("\n--- Sollte eine Textmarke fehlen, bitte die Vorlagendatei nochmal pruefen und ggf. neu einfuegen ---");
            writer.write("\n--- Bitte beachten: Pro Zeile darf nur eine Textmarke existieren ---");
            writer.close();
        }
    }

    private void findBookmarks() {
        // checking header
        for (XWPFHeader header : this.document.getHeaderList()) {
            checkBodyElementsForBookmarks(header.getBodyElements());
        }
        // checking document
        checkBodyElementsForBookmarks(this.document.getBodyElements());
        // checking footer
        for (XWPFFooter footer : this.document.getFooterList()) {
            checkBodyElementsForBookmarks(footer.getBodyElements());
        }
    }

    private void checkBodyElementsForBookmarks(List<IBodyElement> bodyElements) {
        for (IBodyElement element : bodyElements) {
            // check paragraphs
            if (element instanceof XWPFParagraph) {
                checkParagraphForBookmarks((XWPFParagraph) element);
            }
            // check tables
            else if (element instanceof XWPFTable) {
                for (XWPFTableRow row : ((XWPFTable) element).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            checkParagraphForBookmarks(paragraph);
                        }
                    }
                }
            }
        }
    }

    private void checkParagraphForBookmarks(XWPFParagraph paragraph) {
        List<CTBookmark> bookmarkList = paragraph.getCTP().getBookmarkStartList();
        if (bookmarkList != null) {
            for (CTBookmark bookmark : bookmarkList) {
                this.allBookmarks.add(bookmark.getName());
            }
        }
    }

}
