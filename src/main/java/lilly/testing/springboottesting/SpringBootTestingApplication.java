package lilly.testing.springboottesting;

import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class SpringBootTestingApplication {

	public static void main(String[] args) throws IOException {
		String path = "d:\\Projekte\\SpringBootTesting\\generated\\";

		TemplateVerifier bookmarkChecker = new TemplateVerifier();
		//bookmarkChecker.validateBookmarkTemplate(path + "vorlageBookmark.docx");
		bookmarkChecker.validateVariableTemplate(path + "VorlageNeu.docx");

		DocCreator docCreator = new DocCreator();
		int ascii = 8194;
		char asciiAsChar = (char) 8194;
		//System.out.println(ascii + " as Char: [" + asciiAsChar + "]");
		docCreator.loadDocument(path + "VorlageNeu.docx");

		POIDocCreator poiDocCreator = new POIDocCreator();
		//poiDocCreator.Test(path + "data.json");

		//SpringApplication.run(WebApplicationApplication.class, args);
	}

}
