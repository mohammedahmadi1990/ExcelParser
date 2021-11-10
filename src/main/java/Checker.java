import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.util.ArrayList;
import java.util.Formatter;
import java.util.List;
import java.util.Locale;

public class Checker {

    String excelFile;
    String wordFile;
    String outPutWord;
    ArrayList<String[]> data;
    FileInputStream fis;
    Workbook workbook;
    Sheet sheet01;
    Sheet sheet02;
    XWPFDocument inputDoc;

    public Checker(String excelFile, String wordFile, String outPutWord) {
        this.excelFile = excelFile;
        this.wordFile = wordFile;
        this.outPutWord = outPutWord;
    }

    // reads excel tables
    public void read() {
        try {
            fis = new FileInputStream(excelFile);
            workbook = new XSSFWorkbook(fis);
            sheet01 = workbook.getSheetAt(0);
            sheet02 = workbook.getSheetAt(1);

            // Read Cell values in two sheets
            String[] fc_s01 = new String[17];
            String[] fc_s02 = new String[16];
            data = new ArrayList<>();

            Row lastRow = sheet01.getRow(89);
            for (int c = 1; c <= fc_s01.length; c++) {
                Cell cell = lastRow.getCell(c);
                if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
                    Double dd = cell.getNumericCellValue();
                    BigDecimal bd = new BigDecimal(dd).setScale(1, RoundingMode.HALF_EVEN);
                    dd = bd.doubleValue();
                    StringBuilder sb = new StringBuilder();
                    Formatter formatter = new Formatter(sb, Locale.US);
                    formatter.format("%(,.1f", dd);
                    fc_s01[c - 1] = sb.toString();
                } else {
                    fc_s01[c - 1] = cell.getStringCellValue();
                }
            }
            data.add(fc_s01);

            lastRow = sheet02.getRow(89);
            for (int c = 1; c <= fc_s02.length; c++) {
                Cell cell = lastRow.getCell(c);
                if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
                    Double dd = cell.getNumericCellValue();
                    BigDecimal bd = new BigDecimal(dd).setScale(1, RoundingMode.HALF_EVEN);
                    dd = bd.doubleValue();
                    StringBuilder sb = new StringBuilder();
                    Formatter formatter = new Formatter(sb, Locale.US);
                    formatter.format("%(,.1f", dd);
                    fc_s02[c - 1] = sb.toString();
                } else {
                    fc_s02[c - 1] = cell.getStringCellValue();
                }
            }
            data.add(fc_s02);
            fis.close();
            workbook.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    // reads word tables
    public void run() {
        try {
            fis = new FileInputStream(wordFile);
            inputDoc = new XWPFDocument(OPCPackage.open(fis));
            List<XWPFTable> tableList = inputDoc.getTables();
            XWPFTable tables[] = new XWPFTable[2];
            tables[0] = tableList.get(tableList.size() - 2); //4
            tables[1] = tableList.get(tableList.size() - 1); //3
            compareRows(sheet01.getRow(89), tables[0].getRow(tables[0].getRows().size() - 1));
            compareRows(sheet02.getRow(89), tables[1].getRow(tables[1].getRows().size() - 1));

            FileOutputStream fout = new FileOutputStream(outPutWord);
            inputDoc.write(fout);
            fout.close();
            inputDoc.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    // Compare two rows of excel and word documents directly
    public boolean compareRows(Row excelRow, XWPFTableRow wordRow) {
        System.out.println(wordRow.getCell(0));
        for (int c = 2; c < wordRow.getTableCells().size(); c++) {
            Cell cell = excelRow.getCell(c);
            BigDecimal bd = null;
            Double dd = 0.0;
            StringBuilder sb = null;
            if (cell.getCellType() == cell.CELL_TYPE_NUMERIC) {
                dd = cell.getNumericCellValue();
                bd = new BigDecimal(dd).setScale(1, RoundingMode.HALF_EVEN);
                dd = bd.doubleValue();
                sb = new StringBuilder();
                Formatter formatter = new Formatter(sb, Locale.US);
                formatter.format("%(,.1f", dd);
            } else {
                sb = new StringBuilder(cell.getStringCellValue());
            }
            if (!sb.toString().equals(wordRow.getCell(c - 1).getText())) {
                for (XWPFParagraph p : wordRow.getCell(c - 1).getParagraphs()) {
                    for (XWPFRun r : p.getRuns()) {
                        String text = r.getText(0);
                        if (text != null) {
                            r.setText(sb.toString(),0);
                        }
                    }
                }

                System.out.println("Cell #" + (c-1) + " is updated based on Excel value.");
            }

            System.out.println("Word: " + wordRow.getCell(c - 1).getText());
            System.out.println("Excel: " + sb.toString());
        }
        return true;
    }
}
