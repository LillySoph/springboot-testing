package lilly.testing.springboottesting;

import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class SpringBootTestingApplication {

	public static void main(String[] args) throws IOException {
		String path = "d:\\Projekte\\SpringBootTesting\\generated\\";
		String templatePath = path + "Vorlage.docx";
		String docxPath = path + "Bericht.docx";
		String pdfPath_opensagres = path + "BerichtAlsPDF_opensagres.pdf";
		String pdfPath_itext = path + "BerichtAlsPDF_itext.pdf";

		DocCreator docCreator = new DocCreator();
		//docCreator.fillTemplate(templatePath, docxPath);
		docCreator.opensagres_convertDocxToPdf(docxPath, pdfPath_opensagres);
		docCreator.itext_convertDocxToPdf(docxPath, pdfPath_itext);
		//SpringApplication.run(WebApplicationApplication.class, args);
	}

}
