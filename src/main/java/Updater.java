import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.math.RoundingMode;
import java.text.NumberFormat;
import java.util.ArrayList;
import java.util.Formatter;
import java.util.List;
import java.util.Locale;

public class Updater {

    String excelFile;
    String wordFile;
    String outPutWord;
    ArrayList<String[]> data;
    FileInputStream fis;
    Workbook workbook;
    Sheet sheet01;
    Sheet sheet02;

    public Updater(String excelFile, String wordFile, String outPutWord) {
        this.excelFile = excelFile;
        this.wordFile = wordFile;
        this.outPutWord = outPutWord;
    }

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

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public void populate() {
        try {
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
                for (int c = 1; c < row.getTableCells().size(); c++) {
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

            fout.close();
            inputDoc.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }


}
