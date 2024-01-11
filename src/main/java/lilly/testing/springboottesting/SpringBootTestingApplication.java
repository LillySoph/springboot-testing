package lilly.testing.springboottesting;

import org.springframework.boot.autoconfigure.SpringBootApplication;

import java.io.IOException;

@SpringBootApplication
public class SpringBootTestingApplication {

	public static void main(String[] args) throws IOException {
		String path = "d:\\Projekte\\SpringBootTesting\\generated\\";
		DocCreator docCreator = new DocCreator();
		docCreator.fillTemplate(path + "Vorlage.docx");

		//SpringApplication.run(WebApplicationApplication.class, args);
	}

}
