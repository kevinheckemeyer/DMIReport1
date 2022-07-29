import java.io.BufferedReader;
import java.io.FileReader;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class DMIReportReaderMain {
    public static final String COMMA_DELIMITER = ",";
     public static void main(String args[]){
         String basePath = "/Users/georgepeter/Downloads/DMI Reports";
         String reportFilePath = basePath+"/july29.csv";
        DMIReportWriter dmiExcelBatchExport = new DMIReportWriter();
        try {
            List<List<String>> records = new ArrayList<>();
           try (BufferedReader br = new BufferedReader(new FileReader(reportFilePath))) {
                String line;
                while ((line = br.readLine()) != null) {
                    String[] values = line.split(COMMA_DELIMITER);
                    records.add(Arrays.asList(values));

                  }
                dmiExcelBatchExport.writeToExcel(records,basePath);
                System.out.println("Export completed successfully...");

            }
        }catch(Exception ex){
            System.out.println("Export Failed...");
            System.out.println(ex);

        }
    }


}
