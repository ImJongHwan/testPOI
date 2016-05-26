package poi.BenchmarkPOI;

import org.apache.poi.ss.usermodel.*;
import poi.Util.CellStylesUtil;
import poi.Constant;
import poi.TestCasePOI.TcSheet;
import poi.TestCasePOI.TcWorkbook;
import poi.Util.FileUtil;
import poi.Util.TcUtil;

import java.util.List;

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
                initVulnerabilityTC(this.tcWorkbook.getTcSheet(benchmarkSheets.getSheetName()), benchmarkSheets.toString());
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

        CreationHelper createHelper = tcWorkbook.getWorkbook().getCreationHelper();

        Cell cell = tempTcSheet.getCell(1, 'A');
        cell.setCellValue("검사날짜");

        cell = tempTcSheet.getCell(1, 'B');
        // Actually this is not 검사날짜.
        cell.setCellValue(tcWorkbook.getWorkbookCreatedTime());

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

        cell = tempTcSheet.mergeCellRegion(3, 5, 'm', 'm');
        cell.setCellValue("Failed Crawling");

        cell = tempTcSheet.mergeCellRegion(3, 5, 'n', 'n');
        cell.setCellValue("Failed TP");

        cell = tempTcSheet.mergeCellRegion(3, 5, 'o', 'o');
        cell.setCellValue("Failed FP");

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
        cell.setCellFormula("COUNTIF(D:D,FALSE)-1");

        cell = tempTcSheet.getCell(6, 'e');
        cell.setCellFormula("COUNTIF(E:E,\"*TRUE*\")");

        cell = tempTcSheet.getCell(6, 'f');
        cell.setCellFormula("COUNTIF(F:F,\"*FALSE*\")");

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

        cell = tempTcSheet.getCell(6, 'o');
        cell.setCellFormula("COUNTA(N:N)-2");

        initVulnerabilitySheetStyle(tempTcSheet);
        cell = tempTcSheet.getCell(1, 'B');
        cell.getCellStyle().setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm"));
    }

    /**
     * initialize vulnerability column width
     *
     * @param sheet sheet to set column width
     */
    private void initVulnerabilityColumnWidth(TcSheet sheet) {
        sheet.getSheet().setDefaultColumnWidth(Constant.DEFAULT_COLUMN_WIDTH);
        sheet.setColumnWidth('b', Constant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('m', Constant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('n', Constant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('o', Constant.BENCHMARK_TEST_CASE_CELL_WIDTH);

        sheet.setColumnWidth('l', 3);
    }

    /**
     * initialize vulnerability sheet style
     *
     * @param sheet sheet to column width
     */
    private void initVulnerabilitySheetStyle(TcSheet sheet) {
        CellStylesUtil cellStyle = new CellStylesUtil(this.tcWorkbook.getWorkbook());
        Cell cell;

        for(char i = 'a'; i <= 'b'; i++){
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i),cellStyle.DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        }

        for(char i = 'c'; i <= 'j'; i++){
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i), cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for(char i = 'm'; i <= 'o'; i++){
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i),cellStyle.DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        }

        sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex("l"), cellStyle.DEFAULT_THICK_RIGHT_LEFT_BG_GRAY);

        sheet.setSameCellStyle(1, 2, 'a', 'k', cellStyle.BOLD_DEFAULT_THICK_TOP_BOTTOM);
        sheet.setSameCellStyle(1, 2, 'm', 'o', cellStyle.BOLD_DEFAULT_THICK_TOP_BOTTOM);

        sheet.setSameCellStyle(6, 6, 'a', 'k', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(6, 6, 'm', 'o', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);
        cell = sheet.getCell(3, 'b');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);
        cell = sheet.getCell(3, 'k');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(3, 6, 'm', 'o', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);

        cell = sheet.getCell(4, 'a');
        cell.setCellStyle(cellStyle.BOLD_CENTER_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(4, 4, 'g', 'h', cellStyle.BOLD_CENTER_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(4, 4, 'i', 'j', cellStyle.BOLD_CENTER_MIDDLE_RIGHT_LEFT);

        sheet.setSameCellStyle(3, 4, 'c', 'd', cellStyle.BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(3, 4, 'e', 'f', cellStyle.BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(3, 3, 'g', 'j', cellStyle.BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT);

        cell = sheet.getCell(5, 'c');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT);
        cell = sheet.getCell(5, 'e');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT);
        cell = sheet.getCell(5, 'g');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT);
        cell = sheet.getCell(5, 'i');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT);

        cell = sheet.getCell(5, 'a');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT);
        cell = sheet.getCell(5, 'd');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT);
        cell = sheet.getCell(5, 'f');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT);
        cell = sheet.getCell(5, 'h');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT);
        cell = sheet.getCell(5, 'j');
        cell.setCellStyle(cellStyle.BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT);
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

    /**
     * init vulnerability test cases
     *
     * @param tcSheet tcSheet to write
     */
    private void initVulnerabilityTC(TcSheet tcSheet, String vulnerabilityShortName){
        List<String> tpTcList = FileUtil.readFile(Constant.TC_BENCHMARK_PATH + vulnerabilityShortName + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION);
        int tpRowNum = TcUtil.writeDownListInSheet(tpTcList, tcSheet, 7, 'b');
        tcSheet.setSameValueMultiCells(7, tpRowNum - 1, 'c','c', true);
        tcSheet.setTpTcEndRowNum(tpRowNum - 1);

        List<String> fpTcList = FileUtil.readFile(Constant.TC_BENCHMARK_PATH + vulnerabilityShortName + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION);
        int fpRowNum = TcUtil.writeDownListInSheet(fpTcList, tcSheet, tpRowNum, 'b');
        tcSheet.setSameValueMultiCells(tpRowNum, fpRowNum -1, 'd', 'd', false);
        tcSheet.setFpTcEndRowNum(fpRowNum - 1);


        for(int i = 7; i < fpRowNum; i++){
            String formula = "IF(EXACT(F" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            Cell cell = tcSheet.getCell(i,'e');
            cell.setCellFormula(formula);
        }

        for(int i = 7; i < fpRowNum; i++){
            String formula = "IF(COUNTIF($M:$M,B" + i + ") > 0, \"FALSE\", \"\")";
            tcSheet.getCell(i,'f').setCellFormula(formula);
        }

        for(int i = 7; i < tpRowNum; i++){
            String formula = "IF(EXACT(H" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            tcSheet.getCell(i,'g').setCellFormula(formula);
        }

        for(int i = 7; i < tpRowNum; i++){
            String formula = "IF(COUNTIF($N:$N,B" + i + ") > 0, \"FALSE\", \"\")";
            tcSheet.getCell(i,'h').setCellFormula(formula);
        }

        for(int i = tpRowNum; i < fpRowNum; i++){
            String formula = "IF(EXACT(J" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            tcSheet.getCell(i,'i').setCellFormula(formula);
        }

        for(int i = tpRowNum; i < fpRowNum; i++){
            String formula = "IF(COUNTIF($O:$O,B" + i + ") > 0, \"FALSE\", \"\")";
            tcSheet.getCell(i,'j').setCellFormula(formula);
        }
    }

    private void initTotalColumnWidth(TcSheet sheet) {
        sheet.setColumnWidth('a', 30);

        for (char i = 'b'; i < 'k'; i++) {
            sheet.setColumnWidth(i, 10);
        }
    }

    private void initTotalSheetStyle(TcSheet sheet) {
        sheet.setSameCellStyle(1, 1, 'a', 'j', CellStylesUtil.getBoldCenterStyle(this.tcWorkbook.getWorkbook()));
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