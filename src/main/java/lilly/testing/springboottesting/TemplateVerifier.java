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
    List<String> allBookmarks = new ArrayList<String>();
    private XWPFDocument document;
    public void verifyTemplate(String fileName) throws IOException {
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
            writer.write("--- Diese Textmarken wurden in der Datei '" + fileName + "' gefunden ---");
            for (String bookmark : this.allBookmarks) {
                writer.write("\n" + bookmark);
            }
            writer.write("\n--- DATEIENDE ---");
            writer.close();
        }
    }

    private void findBookmarks() {
        // checking Header
        for (XWPFHeader header : this.document.getHeaderList()) {
            checkBodyElements(header.getBodyElements());
        }
        // checking Document
        checkBodyElements(this.document.getBodyElements());
        // checking Footer
        for (XWPFFooter footer : this.document.getFooterList()) {
            checkBodyElements(footer.getBodyElements());
        }
    }

    private void checkBodyElements(List<IBodyElement> bodyElements) {
        for (IBodyElement element : bodyElements) {
            if (element instanceof XWPFParagraph) {
                checkParagraph((XWPFParagraph) element);
            }
            else if (element instanceof XWPFTable) {
                for (XWPFTableRow row : ((XWPFTable) element).getRows()) {
                    for (XWPFTableCell cell : row.getTableCells()) {
                        for (XWPFParagraph paragraph : cell.getParagraphs()) {
                            checkParagraph(paragraph);
                        }
                    }
                }
            }
        }
    }

    private void checkParagraph(XWPFParagraph paragraph) {
        List<CTBookmark> bookmarkList = paragraph.getCTP().getBookmarkStartList();
        if (bookmarkList != null) {
            for (CTBookmark bookmark : bookmarkList) {
                this.allBookmarks.add(bookmark.getName());
            }
        }
    }

}
