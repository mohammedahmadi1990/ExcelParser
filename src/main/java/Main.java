import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.List;

public class Main {

    public static void main(String[] args) {
        try {
            // ********** PART I ********** //

            // Read Word-file
            String excelFile = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\BMS-Excel-Data.xlsx";
            String wordFile = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\MSB_241_en_test.docx";
            String outPutWord = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\result.docx";
            FileInputStream fis = new FileInputStream(excelFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet01 = workbook.getSheetAt(0);
            Sheet sheet02 = workbook.getSheetAt(1);

            // Read Cell values in two sheets
            String[] fc_s01 = new String[17];
            String[] fc_s02 = new String[16];
            ArrayList<String[]> data = new ArrayList<>();

            Row lastRow = sheet01.getRow(89);
            for (int c = 1; c <= fc_s01.length; c++) {
                Cell cell = lastRow.getCell(c);
                fc_s01[c - 1] = cell.toString();
            }
            data.add(fc_s01);

            lastRow = sheet02.getRow(89);
            for (int c = 1; c <= fc_s02.length; c++) {
                Cell cell = lastRow.getCell(c);
                fc_s02[c - 1] = cell.toString();
            }
            data.add(fc_s02);

            // Populate into the tables
            fis = new FileInputStream(wordFile);
            XWPFDocument inputDoc = new XWPFDocument(OPCPackage.open(fis));
            List<XWPFTable> tableList;
            tableList = inputDoc.getTables();
            XWPFTable tables[] = new XWPFTable[2];
            tables[0] = tableList.get(tableList.size() - 2);
            tables[1] = tableList.get(tableList.size() - 1);

            int t = 0;
            for (XWPFTable table :
                    tables) {
                XWPFTableRow row = table.getRow(table.getRows().size() - 1);
                XWPFTableRow sampleRow = table.getRow(table.getRows().size() - 2);
                for (int c = 1; c < data.get(t).length; c++) {
                    XWPFTableCell cell = row.getCell(c);
                    XWPFRun run = cell.addParagraph().createRun();
                    run.setText(data.get(t)[c]);
                    run.setFontSize(sampleRow.getCell(1).getParagraphs().get(0).getRuns().get(0).getFontSize());
                    run.setFontFamily(sampleRow.getCell(1).getParagraphs().get(0).getRuns().get(0).getFontFamily());
                }
                t = t + 1;
            }

            // Update Changes in the Word document
            FileOutputStream fout = new FileOutputStream(outPutWord);
            inputDoc.write(fout);

            // ********** PART II ********** //
            boolean table01Status = compareRows(sheet01.getRow(89),tables[0].getRow(tables[0].getRows().size() - 1));
            boolean table02Status = compareRows(sheet02.getRow(89),tables[1].getRow(tables[1].getRows().size() - 1));
            if(table01Status && table02Status) {
                System.out.println("\n [ Cell values are copied successfully! ] ");
            }else{
                System.out.println("Rows are different!");
            }


            fout.close();
            inputDoc.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    // Compare two rows of excel and word documents directly
    public static boolean compareRows(Row excelRow, XWPFTableRow wordRow){
        for (int c = 2; c < wordRow.getTableCells().size() ; c++) {
            try {
                if (Double.parseDouble(excelRow.getCell(c).toString()) != Double.parseDouble(wordRow.getCell(c - 1).getText())) {
                    return false;
                }
            }catch (Exception e){
                if (!excelRow.getCell(c).toString().equals(wordRow.getCell(c - 1).getText())) {
                    return false;
                }
            }
            System.out.println("Excel: " + wordRow.getCell(c - 1).getText());
            System.out.println("Word: " + excelRow.getCell(c).toString());
        }
        return true;
    }

}
