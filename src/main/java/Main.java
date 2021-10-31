import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;

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
            String outPutWord = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\result.docx";
            FileInputStream fis = new FileInputStream(excelFile);
            Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet01 = workbook.getSheetAt(0);

            // Fields
            int startRow = 8 - 1;
            int endRow = 90;
            int startColumn = 2 - 1;
            int endColumn = 18;
            ArrayList<Integer> loc = new ArrayList<Integer>();
            String[][] data = new String[endRow - startRow][endColumn - startColumn];
            for (int r = startRow; r < endRow; r++) {
                Row row = sheet01.getRow(r);
                for (int c = startColumn; c < endColumn; c++) {
                    Cell cell = row.getCell(c);
                    data[r - startRow][c - startColumn] = cell.toString();
                }
            }

            // Filling Record objects to an Array
            ArrayList<Record> records = new ArrayList<Record>();
            Month month = new Month();
            int tempYear = 0;
            for (int r = 0; r < endRow - startRow; r++) {
                //
                Record record = new Record();

                //set month and year
                if (data[r][0].matches(".*\\d.*")) {
                    String[] str = data[r][0].split(" ");
                    int year = 0;
                    if (str[0].matches(".*\\d.*")) {
                        year = Integer.parseInt(str[0]);
                        record.setYear(year);
                        record.setMonth(str[1].toLowerCase());
                    } else {
                        year = Integer.parseInt(str[1]);
                        record.setMonth(str[0].toLowerCase());
                        record.setYear(year);
                    }
                    tempYear = year;
                    for (int i = 0; i < loc.size(); i++) {
                        records.get(loc.get(i)).setYear(year);
                    }
                    loc = new ArrayList<Integer>();
                } else {
                    loc.add(r);
                    record.setMonth(data[r][0].toLowerCase());
                }
                records.add(record);

                // Set Data
                String[] dd = new String[endColumn - 1];
                for (int c = 1; c < endColumn - 1; c++) {
                    dd[c] = data[r][c];
                }
                record.setData(dd);

                // Check again for years
                if (records.get(r).getYear() == 0) {
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
            XWPFTable table01 = tableList.get(tableList.size() - 2);
            XWPFTable table02 = tableList.get(tableList.size() - 1);

            startRow = 4;
            endRow = table01.getRows().size() - 3 - 1;
            startColumn = 0;
            endColumn = 16;
            int columnCount = endColumn - startColumn + 1;
            Month mnth = new Month();
            for (int r = startRow; r < table01.getRows().size() - 3 - 1; r++) {
                XWPFTableRow row = table01.getRow(r);
                int year = 0;
                String monthName = "";
                int monthInt = 0;
                if (table01.getRow(r).getCell(0).getText().trim().contains(" ")) {
                    year = Integer.parseInt(table01.getRow(r).getCell(0).getText().trim().split(" ")[0]);
                    monthName = table01.getRow(r).getCell(0).getText().trim().toLowerCase().split(" ")[1];
                } else {
                    try {
                        year = Integer.parseInt(table01.getRow(r).getCell(0).getText().trim());
                    } catch (Exception e) {
                        monthName = table01.getRow(r).getCell(0).getText().trim().toLowerCase();
                    }
                }
                if (year == 0) {
                    year = tempYear;
                } else {
                    tempYear = year;
                }
                System.out.println("");

                // if no month name add full year from excel

                for (int i = 0; i < records.size(); i++) {
                    monthInt = mnth.AlbanianMonthToNum(records.get(i).getMonth());
                    int monthInt2 = mnth.EnglishMonthToNum(monthName);
                    if (records.get(i).getYear() == year && monthInt == monthInt2) {
                        for (int c = startColumn; c < columnCount; c++) {
                            row.getCell(c).getParagraphs().get(0).getRuns().get(0).setText(records.get(i).getData()[c], 0);
                        }
//                    } else if (records.get(i).getYear() == year && monthInt <= monthInt2) {
//                        // add new cells to row [prior]
//                        XWPFTableRow newRow = row;
//                        for (int c = startColumn; c < columnCount; c++) {
//                            newRow.getCell(c).getParagraphs().get(0).getRuns().get(0).setText(records.get(i).getData()[c], 0);
//                        }
//                        table01.addRow(newRow,r-1);
//                    } else if (records.get(i).getYear() == year && monthInt >= monthInt2) {
//                        // add new cells to row [after]
//                        XWPFTableRow newRow = row;
//                        for (int c = startColumn; c < columnCount; c++) {
//                            newRow.getCell(c).getParagraphs().get(0).getRuns().get(0).setText(records.get(i).getData()[c], 0);
//                        }
//                        table01.addRow(newRow,r);
//                    } else if (records.get(i).getYear() <= year && monthInt <= monthInt2) {
//                        // add new cells to row [before]
//                        XWPFTableRow newRow = row;
//                        for (int c = startColumn; c < columnCount; c++) {
//                            newRow.getCell(c).getParagraphs().get(0).getRuns().get(0).setText(records.get(i).getData()[c], 0);
//                        }
//                        table01.addRow(newRow,r-1);
//                    } else if (records.get(i).getYear() >= year && monthInt >= monthInt2) {
//                        // add new cells to row [after]
//                        XWPFTableRow newRow = row;
//                        for (int c = startColumn; c < columnCount; c++) {
//                            newRow.getCell(c).getParagraphs().get(0).getRuns().get(0).setText(records.get(i).getData()[c], 0);
//                        }
//                        table01.addRow(newRow,r);
                    }

                }
                if(mnth.EnglishMonthToNum(monthName)==-1){
                    // Delete Row
                    table01.removeRow(r);
                }
            }

            System.out.println("OK");

            // Update Changes in the Word document
            FileOutputStream fout = new FileOutputStream(outPutWord);
            inputDoc.write(fout);
            fout.close();
            inputDoc.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }


}
