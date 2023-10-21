package ExcelReader;
import org.apache.http.HttpResponse;
import org.apache.http.client.HttpClient;
import org.apache.http.client.methods.HttpGet;
import org.apache.http.impl.client.HttpClients;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;

public class JiraToExcel {
    public static void main(String[] args) {
        // Jira API URL for your Jira instance
        String jiraBaseUrl = "https://your-jira-instance";
        String jqlQuery = "project = YOUR_PROJECT";

        try {
            // Authenticate and make the Jira API request
            HttpClient httpClient = HttpClients.createDefault();
            HttpGet request = new HttpGet(jiraBaseUrl + "/rest/api/2/search?jql=" + jqlQuery);
            // Set Basic Authentication credentials if needed

            HttpResponse response = httpClient.execute(request);

            // Parse the JSON response to extract issue data

            // Create a new Excel workbook
            Workbook workbook = new XSSFWorkbook();
            Sheet sheet = workbook.createSheet("Jira Issues");

            // Add headers to the spreadsheet
            Row headerRow = sheet.createRow(0);
            String[] headers = {"Issue Key", "Summary", "Description", "Status", "Assignee"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
            }

            // Iterate through the Jira issues and add them to the spreadsheet
            // Replace this with code to populate issue data from your JSON response

            // Save the Excel file
            FileOutputStream outputStream = new FileOutputStream("JiraIssues.xlsx");
            workbook.write(outputStream);
            outputStream.close();

            System.out.println("Excel file created successfully.");


            
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    private static void iterateOverJson(JsonNode node, String path) {
    if (node.isObject()) {
        node.fields().forEachRemaining(entry -> {
            String key = entry.getKey();
            JsonNode value = entry.getValue();
            String newPath = path.isEmpty() ? key : path + "." + key;

            System.out.println("Key: " + newPath);
            if (value.isObject() || value.isArray()) {
                iterateOverJson(value, newPath);
            } else {
                System.out.println("Value: " + value);
            }
        });
    } else if (node.isArray()) {
        for (int i = 0; i < node.size(); i++) {
            JsonNode arrayElement = node.get(i);
            iterateOverJson(arrayElement, path + "[" + i + "]");
        }
    }
}
