package poi.BenchmarkPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.XSSFFont;
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
                this.tcWorkbook.getTcSheet(benchmarkSheets.getSheetName());
            }
        }
    }

    private void initVulnerabilitySheet(TcSheet tempTcSheet) {
//        TcSheet tempTcSheet = this.tcWorkbook.getTcSheet(vulnerabilitySheetName);
//        XSSFCellStyle cs = tcWorkbook.getWorkbook().createCellStyle();
//        cs.setAlignment(HorizontalAlignment.CENTER);

        CellStyle dateCellStyle = tcWorkbook.getWorkbook().createCellStyle();
        CreationHelper createHelper = tcWorkbook.getWorkbook().getCreationHelper();

        dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd hh:mm"));
        dateCellStyle.setFont();

        XSSFFont titleFont = tcWorkbook.getWorkbook().createFont();
        titleFont.setBold(true);

        CellStyle centerTitleCell = tcWorkbook.getWorkbook().createCellStyle();
        centerTitleCell.setFont(titleFont);
        centerTitleCell.setAlignment(CellStyle.ALIGN_CENTER);

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