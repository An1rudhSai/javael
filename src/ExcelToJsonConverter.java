import com.aspose.cells.Workbook;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.TxtLoadOptions;

public class ExcelToJsonConverter {

    public static void main(String[] args) {
        try {
            // Load the XLS file
            String xlsFilePath = "/Users/anirudhsai/Downloads/file_example_XLS_100.xls"; // Provide the actual path
            Workbook workbook = new Workbook(xlsFilePath);

            // Set the options for saving as JSON
            TxtLoadOptions loadOptions = new TxtLoadOptions();
            loadOptions.setSeparator('\t'); // Use tab as separator for JSON format
            loadOptions.setConvertNumericData(true); // Convert numeric data to string

            // Generate a new JSON file name based on the Excel file name
            String excelFileName = xlsFilePath.substring(xlsFilePath.lastIndexOf('/') + 1);
            String jsonFileName = excelFileName.replace(".xls", ".json");

            // Save XLS as JSON in the same directory
            String jsonFilePath = xlsFilePath.replace(excelFileName, jsonFileName);
            workbook.save(jsonFilePath, SaveFormat.JSON);

            System.out.println("XLS file converted to JSON: " + jsonFilePath);
        } catch (Exception e) {
            // Handle the exception
            e.printStackTrace();
        }
    }
}
