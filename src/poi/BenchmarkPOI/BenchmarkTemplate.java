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
     *
     */
    private void init(){
        for(Constant.BenchmarkSheets benchmarkSheets : Constant.BenchmarkSheets.values()){
            initVulnerabilitySheet(this.tcWorkbook.getTcSheet(benchmarkSheets.getSheetName()));
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

        cell = tempTcSheet.getCell(6,'a');
        cell.setCellValue("합계");

//        cell = tempTcSheet.getCell(3,'b');
//        tempTcSheet.getSheet().addMergedRegion(CellRangeAddress);
//
//        CellRangeAddress cra = new CellRangeAddress();
    }

    /**
     * write current state in excel file - ZAP_OWASPBenchmark_Results_[time].xlsx
     */
    public void writeBenchmark(){
        if(tcWorkbook != null){
            tcWorkbook.writeWorkbook(Constant.BENCHMARK_NAME);
        }
    }
}
