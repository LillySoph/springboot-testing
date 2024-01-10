package lilly.testing.springboottesting;

import org.json.simple.JSONArray;
import org.json.simple.JSONObject;
import org.json.simple.parser.*;

import java.io.*;
import java.util.HashMap;
import java.util.Map;

public class POIDocCreator {

    private Map<String, String> variableReplacement;

    public void Test(String fileName) throws IOException {
        System.out.println(readJSONFile(fileName));
    }

    private JSONObject readJSONFile(String fileName) throws IOException {
        // check if file exists
        File file = new File(fileName);
        if (!file.exists()) {
            return null;
            //throw new FileNotFoundException("Die Datei '" + fileName + "' konnte nicht gefunden werden.");
        }
        // parse json file
        try(FileReader fileReader = new FileReader(file)) {
            Object obj = new JSONParser().parse(fileReader);
            JSONObject jsonObject = (JSONObject) obj;
            return jsonObject;//jsonObject;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private void createTestData() {

        this.variableReplacement = new HashMap<String, String>();
        this.variableReplacement.put("Datum", "01.12.3005");
        this.variableReplacement.put("Betreff", "Das ist der Betreff");
        this.variableReplacement.put("TgbNr", "12345");
        this.variableReplacement.put("hier1", "Das wird ziemlich lecker.");
        this.variableReplacement.put("hier2", "LeckerLeckerLecker");
        this.variableReplacement.put("Bezug", "Wird gebacken für die Abschlussfeier von Cola");
        this.variableReplacement.put("verdeckteErhebung", "LeckerLecker YumYum");
    }

}
