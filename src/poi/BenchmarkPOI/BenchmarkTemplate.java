package poi.BenchmarkPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFFont;
import poi.CellStyles;
import poi.Constant;
import poi.TestCasePOI.TcSheet;
import poi.TestCasePOI.TcWorkbook;

/**
 * Created by Hwan on 2016-05-10.
 */
public class BenchmarkTemplate {
    private TcWorkbook tcWorkbook;

    public BenchmarkTemplate() {
        this.tcWorkbook = new TcWorkbook();

        init();
    }

    /**
     * Set Benchmark workbook templates.
     */
    private void init() {
        for (Constant.BenchmarkSheets benchmarkSheets : Constant.BenchmarkSheets.values()) {
            if (!benchmarkSheets.getSheetName().equals(Constant.BenchmarkSheets.total.getSheetName())) {
                initVulnerabilitySheet(this.tcWorkbook.getTcSheet(benchmarkSheets.getSheetName()));
            } else {
                initTotalSheet(this.tcWorkbook.getTcSheet(benchmarkSheets.getSheetName()));
            }
        }
    }

    /**
     * generate vulnerability sheet template
     *
     * @param tempTcSheet vulnerability sheet
     */
    private void initVulnerabilitySheet(TcSheet tempTcSheet) {
        initVulnerabilityColumnWidth(tempTcSheet);
        initVulnerabilitySheetStyle(tempTcSheet);

        CellStyle dateCellStyle = tcWorkbook.getWorkbook().createCellStyle();
        CreationHelper createHelper = tcWorkbook.getWorkbook().getCreationHelper();

        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm"));

        Cell cell = tempTcSheet.getCell(1, 'A');
        cell.setCellValue("검사날짜");

        cell = tempTcSheet.getCell(1, 'B');
        // Actually this is not 검사날짜.
        cell.setCellValue(tcWorkbook.getWorkbookCreatedTime());
        cell.setCellStyle(dateCellStyle);

        cell = tempTcSheet.getCell(2, 'a');
//        cell.setCellValue(vulnerabilitySheetName);
        cell.setCellValue(tempTcSheet.getSheet().getSheetName());

        cell = tempTcSheet.mergeCellRegion(3, 5, 'b', 'b');
        cell.setCellValue("Test Cases");

        cell = tempTcSheet.mergeCellRegion(3, 4, 'c', 'd');
        cell.setCellValue("Real Vulnerability");

        cell = tempTcSheet.mergeCellRegion(3, 4, 'e', 'f');
        cell.setCellValue("URL Crawl");

        cell = tempTcSheet.mergeCellRegion(3, 3, 'g', 'j');
        cell.setCellValue("Detected");

        cell = tempTcSheet.mergeCellRegion(4, 4, 'g', 'h');
        cell.setCellValue("TRUE Vulnerability");

        cell = tempTcSheet.mergeCellRegion(4, 4, 'i', 'j');
        cell.setCellValue("FALSE Vulnerability");

        cell = tempTcSheet.mergeCellRegion(3, 5, 'k', 'k');
        cell.setCellValue("Description");

        cell = tempTcSheet.mergeCellRegion(1, Constant.MAX_ROW, 'l', 'l');

        cell = tempTcSheet.mergeCellRegion(3, 5, 'm', 'm');
        cell.setCellValue("Failed TP");

        cell = tempTcSheet.mergeCellRegion(3, 5, 'n', 'n');
        cell.setCellValue("Failed TN");

        cell = tempTcSheet.getCell(6, 'a');
        cell.setCellValue("합계");

        cell = tempTcSheet.getCell(5, 'c');
        cell.setCellValue(true);

        cell = tempTcSheet.getCell(5, 'd');
        cell.setCellValue(false);

        cell = tempTcSheet.getCell(5, 'e');
        cell.setCellValue(true);

        cell = tempTcSheet.getCell(5, 'f');
        cell.setCellValue(false);

        cell = tempTcSheet.getCell(5, 'g');
        cell.setCellValue(true);

        cell = tempTcSheet.getCell(5, 'h');
        cell.setCellValue(false);

        cell = tempTcSheet.getCell(5, 'i');
        cell.setCellValue(true);

        cell = tempTcSheet.getCell(5, 'j');
        cell.setCellValue(false);

        cell = tempTcSheet.getCell(6, 'b');
        cell.setCellFormula("COUNTA(B:B)-3");

        cell = tempTcSheet.getCell(6, 'c');
        cell.setCellFormula("COUNTIF(C:C,TRUE)-1");

        cell = tempTcSheet.getCell(6, 'd');
        cell.setCellFormula("COUNTIF(D:D,\"FALSE\")-1");

        cell = tempTcSheet.getCell(6, 'e');
        cell.setCellFormula("COUNTIF(E:E,TRUE)-1");

        cell = tempTcSheet.getCell(6, 'f');
        cell.setCellFormula("COUNTIF(F:F,\"FALSE\")-1");

        cell = tempTcSheet.getCell(6, 'g');
        cell.setCellFormula("COUNTIF(G:G,\"*TRUE*\")-1");

        cell = tempTcSheet.getCell(6, 'h');
        cell.setCellFormula("COUNTIF(H:H,\"*FALSE*\")");

        cell = tempTcSheet.getCell(6, 'i');
        cell.setCellFormula("COUNTIF(I:I,\"*TRUE*\")");

        cell = tempTcSheet.getCell(6, 'j');
        cell.setCellFormula("COUNTIF(J:J,\"*FALSE*\")");

        cell = tempTcSheet.getCell(6, 'k');
        cell.setCellFormula("COUNTA(K:K)-2");

        cell = tempTcSheet.getCell(6, 'm');
        cell.setCellFormula("COUNTA(M:M)-2");

        cell = tempTcSheet.getCell(6, 'n');
        cell.setCellFormula("COUNTA(N:N)-2");
    }

    /**
     * initialize vulnerability column width
     *
     * @param sheet sheet to set column width
     */
    private void initVulnerabilityColumnWidth(TcSheet sheet) {
        sheet.setColumnWidth('b', Constant.BENCHMARK_TEST_CASE_CELL_SIZE);
        sheet.setColumnWidth('m', Constant.BENCHMARK_TEST_CASE_CELL_SIZE);
        sheet.setColumnWidth('n', Constant.BENCHMARK_TEST_CASE_CELL_SIZE);

        sheet.setColumnWidth('l', 3);
    }

    /**
     * initialize vulnerability sheet style
     *
     * @param sheet sheet to column width
     */
    private void initVulnerabilitySheetStyle(TcSheet sheet) {
        sheet.setSameCellStyle(1, 2, 'a', 'b', CellStyles.getBoldStyle(this.tcWorkbook.getWorkbook()));

        sheet.setSameCellStyle(3, 6, 'a', 'n', CellStyles.getBoldCenterStyle(this.tcWorkbook.getWorkbook()));
    }


    /**
     * initialize total sheet
     *
     * @param tempTcSheet total sheet
     */
    private void initTotalSheet(TcSheet tempTcSheet) {

        initTotalColumnWidth(tempTcSheet);
        initTotalSheetStyle(tempTcSheet);

        Cell cell = tempTcSheet.getCell(1, 'A');
        cell.setCellValue("Category");

        cell = tempTcSheet.getCell(1, 'B');
        cell.setCellValue("CWE #");

        cell = tempTcSheet.getCell(1, 'c');
        cell.setCellValue("TP");

        cell = tempTcSheet.getCell(1, 'd');
        cell.setCellValue("FN");

        cell = tempTcSheet.getCell(1, 'e');
        cell.setCellValue("TN");

        cell = tempTcSheet.getCell(1, 'f');
        cell.setCellValue("FP");

        cell = tempTcSheet.getCell(1, 'g');
        cell.setCellValue("Total");

        cell = tempTcSheet.getCell(1, 'h');
        cell.setCellValue("TPR");

        cell = tempTcSheet.getCell(1, 'i');
        cell.setCellValue("FPR");

        cell = tempTcSheet.getCell(1, 'j');
        cell.setCellValue("Score");

        int rowNum = 1;

        for (Constant.BenchmarkSheets vulnerability : Constant.BenchmarkSheets.values()) {
            if (!vulnerability.getSheetName().equals("Total")) {
                rowNum++;
                cell = tempTcSheet.getCell(rowNum, 'a');
                cell.setCellValue(vulnerability.getSheetName());
            }
        }

        cell = tempTcSheet.getCell(13, 'a');
        cell.setCellValue("Totals");

        cell = tempTcSheet.getCell(14, 'a');
        cell.setCellValue("Overall Results");

    }

    private void initTotalColumnWidth(TcSheet sheet){
        sheet.setColumnWidth('a', 30);

        for(char i = 'b'; i < 'k'; i++){
            sheet.setColumnWidth(i, 10);
        }
    }

    private void initTotalSheetStyle(TcSheet sheet){
        sheet.setSameCellStyle(1,1,'a','j',CellStyles.getBoldCenterStyle(this.tcWorkbook.getWorkbook()));
    }

    //todo define default XSSFCEllStyle

    /**
     * write current state in excel file - ZAP_OWASPBenchmark_Results_[time].xlsx
     */
    public void writeBenchmark() {
        if (tcWorkbook != null) {
            tcWorkbook.writeWorkbook(Constant.BENCHMARK_NAME);
        }
    }
}