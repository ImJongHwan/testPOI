package poi;

import poi.BenchmarkPOI.BenchmarkTemplate;

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
