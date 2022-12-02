package epic2sectra;

import java.io.File;
import java.io.IOException;
import java.io.PrintStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbookFactory;

/**
 *
 * @author Geoff
 */
public class ConvertXlsx {

    public static void main(String[] args) throws IOException {

        DateFormat dfDay = new SimpleDateFormat("yyyyMMdd");
        DateFormat dfTimestamp = new SimpleDateFormat("yyyyMMddHHmmss");
        
        String excelFilePath = args[0];
        String password = args[1];
        
        WorkbookFactory.addProvider(new XSSFWorkbookFactory());
        Workbook workbook = WorkbookFactory.create(new File(excelFilePath), password);
        Sheet sheet = workbook.getSheetAt(0);

        System.out.println(String.format("reading from file %s", excelFilePath));
        
        PrintStream csv = new PrintStream(new File(excelFilePath.replace(".xlsx", ".csv")));

        System.out.println(String.format("writing to file   %s", excelFilePath.replace(".xlsx", ".csv")));
        
        csv.println(String.format("\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"",
          "slideBarCode",
          "service",
          "accNo",
          "partId",
          "blockId",
          "slideNo",
          "stain",
          "mrn",
          "empi",
          "dob",
          "lastName",
          "firstName",
          "gender",
          "collectionDt",
          "orderDt"
        ));
        
        Iterator<Row> rowIterator = sheet.iterator();
        Row headerRow = rowIterator.next();
        Map<String, Integer> columnIndexByNameMap = new HashMap<>();
        columnIndexByNameMap.put("Slide Bar Code", Integer.valueOf(headerRow.getFirstCellNum())); // Epic uses "Container" for two different columns
        for(int x = headerRow.getFirstCellNum() + 1; x <= headerRow.getLastCellNum(); x++) {
            if(headerRow.getCell(x) != null) {
                columnIndexByNameMap.put(headerRow.getCell(x).getStringCellValue(), x);
            }
        }
        
        int rowsProcessed = 0;
        int rowsSkipped = 0;
        while(rowIterator.hasNext()) {
            
            Row dataRow = rowIterator.next();

            if(dataRow.getCell(columnIndexByNameMap.get("Slide Bar Code")) == null) {
                rowsSkipped++;
                continue;
            }

            rowsProcessed++;
            
            String gender = dataRow.getCell(columnIndexByNameMap.get("Gender")).getStringCellValue().substring(0, 1);
            if(!("M".equals(gender) || "F".equals(gender))) { gender = "M"; } // for the moment, M is the default gender
            
            csv.println(String.format("\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\",\"%s\"",
              dataRow.getCell(columnIndexByNameMap.get("Slide Bar Code")).getStringCellValue(),
              dataRow.getCell(columnIndexByNameMap.get("Specialty")).getStringCellValue(),
              dataRow.getCell(columnIndexByNameMap.get("Specimen/Case ID")).getStringCellValue(),
              dataRow.getCell(columnIndexByNameMap.get("Container")).getStringCellValue().split(",")[1].trim(),
              dataRow.getCell(columnIndexByNameMap.get("Container")).getStringCellValue().split(",")[2].trim(),
              dataRow.getCell(columnIndexByNameMap.get("Container")).getStringCellValue().split(",")[3].trim(),
              dataRow.getCell(columnIndexByNameMap.get("Task")).getStringCellValue(),
              dataRow.getCell(columnIndexByNameMap.get("MRN")).getStringCellValue(),
              dataRow.getCell(columnIndexByNameMap.get("Patient Enterprise ID")).getStringCellValue(),
              dfDay.format(dataRow.getCell(columnIndexByNameMap.get("Birth Date")).getDateCellValue()),
              dataRow.getCell(columnIndexByNameMap.get("Patient Last Name")).getStringCellValue(),
              dataRow.getCell(columnIndexByNameMap.get("Patient First Name")).getStringCellValue(),
              gender,
              dfTimestamp.format(dataRow.getCell(columnIndexByNameMap.get("Collected")).getDateCellValue()),
              dfTimestamp.format(dataRow.getCell(columnIndexByNameMap.get("Ordered Instant")).getDateCellValue())
            ));

        }
        
        csv.close();
        
        System.out.println(String.format("%5d rows processed", rowsProcessed));
        System.out.println(String.format("%5d rows skipped", rowsSkipped));
        
    }
    
}
