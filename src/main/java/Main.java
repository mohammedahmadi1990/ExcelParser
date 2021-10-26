import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        try {
            // Read Word-file
            //FileInputStream fis = new FileInputStream(args[0]);
            String excelFile = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\BMS-Excel-Data.xlsx";
            String wordFile = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\MSB_241_en_test.docx";
            FileInputStream fis = new FileInputStream(excelFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet01 = workbook.getSheetAt(0);

            // Fields
            int startRow = 8 - 1;
            int endRow = 90;
            int startColumn = 2 - 1;
            int endColumn = 18;
            ArrayList<Integer> loc = new ArrayList<Integer>();
            String[][] data = new String[endRow-startRow][endColumn-startColumn];
            for (int r = startRow; r < endRow; r++) {
                Row row = sheet01.getRow(r);
                for (int c = startColumn; c < endColumn ; c++) {
                    Cell cell = row.getCell(c);
                    data[r-startRow][c-startColumn] = cell.toString();
                }
            }

            // Filling Record objects to an Array
            ArrayList<Record> records = new ArrayList<Record>();
            Month month = new Month();
            int tempYear = 0;
            for (int r = 0; r < endRow-startRow; r++) {
                //
                Record record = new Record();

                //set month and year
                if(data[r][0].matches(".*\\d.*")){
                    String[] str = data[r][0].split(" ");
                    int year = 0;
                    if(str[0].matches(".*\\d.*")){
                        year = Integer.parseInt(str[0]);
                       record.setYear(year);
                       record.setMonth(str[1]);
                    }else{
                        year = Integer.parseInt(str[1]);
                        record.setMonth(str[0]);
                        record.setYear(year);
                    }
                    tempYear = year;
                    for (int i = 0; i < loc.size(); i++) {
                        records.get(loc.get(i)).setYear(year);
                    }
                    loc = new ArrayList<Integer>();
                }else{
                    loc.add(r);
                    record.setMonth(data[r][0]);
                }
                records.add(record);

                // Set Data
                String[] dd = new String[endColumn-1];
                for (int c = 1; c < endColumn-1; c++) {
                       dd[c] = data[r][c];
                }
                record.setData(dd);

                // Check again for years
                if(records.get(r).getYear()==0){
                    for (int i = 0; i < loc.size(); i++) {
                        records.get(loc.get(i)).setYear(tempYear);
                    }
                }
            }

            // Read word-document in order to compare data
            fis = new FileInputStream(wordFile);
            XWPFDocument inputDoc = new XWPFDocument(OPCPackage.open(fis));
            List<XWPFTable> tableList;
            tableList = inputDoc.getTables();
            XWPFTable table01 = tableList.get(tableList.size()-2);
            XWPFTable table02 = tableList.get(tableList.size()-1);

//            for (int r = 4; r < 30; r++) {
//                XWPFTableRow row = table.getRow(r);
//                int year = Integer.parseInt(table.getRow(r).getCell(0).getText().trim());
//                for (int c = 0; c < 3; c++) {
//                    row.getCell(c).setText("45");
//                }
//                table.addRow(row,r);
//                System.out.println(r);
//            }

            System.out.println("OK");

            // Update Changes in the Word document
            FileOutputStream fout = new FileOutputStream(wordFile);
            inputDoc.write(fout);
            fout.close();
            inputDoc.close();
        } catch(Exception ex) {
            ex.printStackTrace();
        }
    }


}
