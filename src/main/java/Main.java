import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;

public class Main {

    public static void main(String[] args) {

        String excelFile = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\table02\\BMS-Excel-Data.xlsx";
        String wordFile = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\table02\\MSB_241_en_test.docx";
        //String outPutWord = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\table01\\MSB_241_en_test.docx";
        String outPutWord = "C:\\Users\\Mohammed\\Desktop\\UPWORK\\result.docx";

//        Updater updater = new Updater(excelFile,wordFile,outPutWord);
//        updater.read();
//        updater.populate();

        Checker checker = new Checker(excelFile,outPutWord,outPutWord);
        checker.read();
        checker.run();

    }
}
