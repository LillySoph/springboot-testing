package lilly.testing.springboottesting;

import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class SpringBootTestingApplication {

	public static void main(String[] args) throws IOException {
		String path = "d:\\Projekte\\SpringBootTesting\\generated\\";
		DocCreator docCreator = new DocCreator();
		int ascii = 8194;
		char asciiAsChar = (char) 8194;
		//System.out.println(ascii + " as Char: [" + asciiAsChar + "]");
		docCreator.loadDocument(path + "vorlageBookmark.docx");

		//SpringApplication.run(WebApplicationApplication.class, args);
	}

}
