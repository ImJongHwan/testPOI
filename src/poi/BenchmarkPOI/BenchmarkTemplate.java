package poi.BenchmarkPOI;

import org.apache.poi.ss.usermodel.Cell;
import poi.Constant;
import poi.TestCasePOI.TcSheet;
import poi.TestCasePOI.TcWorkbook;

/**
 * Created by Hwan on 2016-05-10.
 */
public class BenchmarkTemplate {
    private TcWorkbook tcWorkbook;

    public BenchmarkTemplate(){
        this.tcWorkbook = new TcWorkbook();

        init();
    }

    /**
     * Set Benchmark workbook templates.
     */
    private void init(){
        for(Constant.BenchmarkSheets benchmarkSheets : Constant.BenchmarkSheets.values()){
            if(!benchmarkSheets.getSheetName().equals(Constant.BenchmarkSheets.total)) {
                initVulnerabilitySheet(this.tcWorkbook.getTcSheet(benchmarkSheets.getSheetName()));
            }
        }
    }

    private void initVulnerabilitySheet(TcSheet tempTcSheet){
//        TcSheet tempTcSheet = this.tcWorkbook.getTcSheet(vulnerabilitySheetName);
//        XSSFCellStyle cs = tcWorkbook.getWorkbook().createCellStyle();
//        cs.setAlignment(HorizontalAlignment.CENTER);

        Cell cell = tempTcSheet.getCell(1,'A');
        cell.setCellValue("검사날짜");

        cell = tempTcSheet.getCell(1, 'B');
        // Actually this is not 검사날짜.
        cell.setCellValue(tcWorkbook.getWorkbookCreatedTime());

        cell = tempTcSheet.getCell(2, 'a');
//        cell.setCellValue(vulnerabilitySheetName);
        cell.setCellValue(tempTcSheet.getSheet().getSheetName());

        cell = tempTcSheet.mergeCellRegion(3,5,'b','b');
        cell.setCellValue("Test Cases");

        cell = tempTcSheet.mergeCellRegion(3,4,'c','d');
        cell.setCellValue("Real Vulnerability");

        cell = tempTcSheet.mergeCellRegion(3,4,'e','f');
        cell.setCellValue("URL Crawl");

        cell = tempTcSheet.mergeCellRegion(3,3,'g','j');
        cell.setCellValue("Detected");

        cell = tempTcSheet.mergeCellRegion(4,4,'g','h');
        cell.setCellValue("TRUE Vulnerability");

        cell = tempTcSheet.mergeCellRegion(4,4,'i','j');
        cell.setCellValue("FALSE Vulnerability");

        cell = tempTcSheet.mergeCellRegion(3,5,'k','k');
        cell.setCellValue("Description");

        cell = tempTcSheet.mergeCellRegion(1,Constant.MAX_ROW, 'l','l');

        cell = tempTcSheet.mergeCellRegion(3,5,'m','m');
        cell.setCellValue("Failed TP");

        cell = tempTcSheet.mergeCellRegion(3,5,'n','n');
        cell.setCellValue("Failed TN");

        cell = tempTcSheet.getCell(6,'a');
        cell.setCellValue("합계");

        cell = tempTcSheet.getCell(5,'c');
        cell.setCellValue("TRUE");

        cell = tempTcSheet.getCell(5,'d');
        cell.setCellValue("FALSE");

        cell = tempTcSheet.getCell(5,'e');
        cell.setCellValue("TRUE");

        cell = tempTcSheet.getCell(5,'f');
        cell.setCellValue("FALSE");

        cell = tempTcSheet.getCell(5,'g');
        cell.setCellValue("TRUE");

        cell = tempTcSheet.getCell(5,'h');
        cell.setCellValue("FALSE");

        cell = tempTcSheet.getCell(5,'i');
        cell.setCellValue("TRUE");

        cell = tempTcSheet.getCell(5,'j');
        cell.setCellValue("FALSE");

        cell = tempTcSheet.getCell(6,'b');
        cell.setCellValue("=COUNTA(B:B)-3");

        cell = tempTcSheet.getCell(6, 'c');
        cell.setCellValue("=COUNTIF(C:C,TRUE)-1");

        cell = tempTcSheet.getCell(6,'d');
        cell.setCellValue("=COUNTIF(D:D,\"FALSE\")-1");

        cell = tempTcSheet.getCell(6, 'e');
        cell.setCellValue("=COUNTIF(E:E,TRUE)-1");

        cell = tempTcSheet.getCell(6, 'f');
        cell.setCellValue("=COUNTIF(F:F,\"FALSE\")-1");

        cell = tempTcSheet.getCell(6,'g');
        cell.setCellValue("=COUNTIF(G:G,\"*TRUE*\")-1");

        cell = tempTcSheet.getCell(6,'h');
        cell.setCellValue("=COUNTIF(H:H,\"*FALSE*\")");

        cell = tempTcSheet.getCell(6,'i');
        cell.setCellValue("=COUNTIF(I:I,\"*TRUE*\")");

        cell = tempTcSheet.getCell(6,'j');
        cell.setCellValue("=COUNTIF(J:J,\"*FALSE*\")");

        cell = tempTcSheet.getCell(6,'k');
        cell.setCellValue("=COUNTA(K:K)-2");

        cell = tempTcSheet.getCell(6, 'm');
        cell.setCellValue("=COUNTA(M:M)-2");

        cell = tempTcSheet.getCell(6, 'n');
        cell.setCellValue("=COUNTA(N:N)-2");
    }

    //todo define default XSSFCEllStyle

    /**
     * write current state in excel file - ZAP_OWASPBenchmark_Results_[time].xlsx
     */
    public void writeBenchmark(){
        if(tcWorkbook != null){
            tcWorkbook.writeWorkbook(Constant.BENCHMARK_NAME);
        }
    }
}
