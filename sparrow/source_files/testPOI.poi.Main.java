package poi;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import poi.BenchmarkPOI.BenchmarkTemplate;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

public class Main {

    public static void main(String[] args) {
//        XSSFWorkbook workbook = new XSSFWorkbook();
//        try {
//            FileOutputStream out = new FileOutputStream(new File("./createworkbook.xlsx"));
//            XSSFSheet xs = workbook.createSheet();
//            XSSFRow xr = xs.createRow(0);
//            XSSFCell xc = xr.createCell(0);
//            xc.setCellValue("=COUNTA(B:B)-3");
//
//            workbook.write(out);
//            out.close();
//        } catch (FileNotFoundException e) {
//            e.printStackTrace();
//        } catch (IOException e) {
//            e.printStackTrace();
//        }

//        System.out.println("createworkbool.xlsx written successfully");
        BenchmarkTemplate bt = new BenchmarkTemplate();
        bt.writeBenchmark();
    }
}
