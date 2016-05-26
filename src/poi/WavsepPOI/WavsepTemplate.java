package poi.WavsepPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CreationHelper;
import poi.Constant;
import poi.TestCasePOI.TcSheet;
import poi.TestCasePOI.TcWorkbook;
import poi.Util.CellStylesUtil;
import poi.Util.FileUtil;
import poi.Util.TcUtil;

import java.util.List;

/**
 * Created by Hwan on 2016-05-26.
 */
public class WavsepTemplate {
    private TcWorkbook tcWorkbook;
    private CellStylesUtil cellStyle;

    public WavsepTemplate() {
        this.tcWorkbook = new TcWorkbook();
        cellStyle = new CellStylesUtil(this.tcWorkbook.getWorkbook());
        init();
        this.tcWorkbook.getWorkbook().setForceFormulaRecalculation(true);
    }

    /**
     * Set wavsep workbook templates.
     */
    private void init() {
        for (Constant.WavsepSheets wavsepSheet : Constant.WavsepSheets.values()) {
            if (!wavsepSheet.getSheetName().equals(Constant.WavsepSheets.total.getSheetName())) {
                initVulnerabilitySheet(this.tcWorkbook.getTcSheet(wavsepSheet.getSheetName()));
                initVulnerabilityTC(this.tcWorkbook.getTcSheet(wavsepSheet.getSheetName()), wavsepSheet.toString());
            } else {
                initTotalSheet(this.tcWorkbook.getTcSheet(wavsepSheet.getSheetName()));
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

        cell = tempTcSheet.mergeCellRegion(3, 4, 'c', 'e');
        cell.setCellValue("Type");

        cell = tempTcSheet.mergeCellRegion(3, 4, 'f', 'g');
        cell.setCellValue("URL Crawl");

        cell = tempTcSheet.mergeCellRegion(3, 3, 'h', 'm');
        cell.setCellValue("Detected");

        cell = tempTcSheet.mergeCellRegion(4, 4, 'h', 'i');
        cell.setCellValue("TRUE Positive");

        cell = tempTcSheet.mergeCellRegion(4, 4, 'j', 'k');
        cell.setCellValue("FALSE Positive");

        cell = tempTcSheet.mergeCellRegion(4, 4, 'l', 'm');
        cell.setCellValue("Experimental");

        cell = tempTcSheet.mergeCellRegion(3, 5, 'n', 'n');
        cell.setCellValue("Description");

        cell = tempTcSheet.mergeCellRegion(3, 5, 'p', 'p');
        cell.setCellValue("Failed Crawling");

        cell = tempTcSheet.mergeCellRegion(3, 5, 'q', 'q');
        cell.setCellValue("Failed TP");

        cell = tempTcSheet.mergeCellRegion(3, 5, 'r', 'r');
        cell.setCellValue("Failed FP");

        cell = tempTcSheet.mergeCellRegion(3, 5, 's', 's');
        cell.setCellValue("Failed EX");

        cell = tempTcSheet.getCell(6, 'a');
        cell.setCellValue("합계");

        cell = tempTcSheet.getCell(5, 'c');
        cell.setCellValue("TP");

        cell = tempTcSheet.getCell(5, 'd');
        cell.setCellValue("FP");

        cell = tempTcSheet.getCell(5, 'e');
        cell.setCellValue("EX");

        cell = tempTcSheet.getCell(5, 'f');
        cell.setCellValue(true);

        cell = tempTcSheet.getCell(5, 'g');
        cell.setCellValue(false);

        cell = tempTcSheet.getCell(5, 'h');
        cell.setCellValue(true);

        cell = tempTcSheet.getCell(5, 'i');
        cell.setCellValue(false);

        cell = tempTcSheet.getCell(5, 'j');
        cell.setCellValue(true);

        cell = tempTcSheet.getCell(5, 'k');
        cell.setCellValue(false);

        cell = tempTcSheet.getCell(5, 'l');
        cell.setCellValue(true);

        cell = tempTcSheet.getCell(5, 'm');
        cell.setCellValue(false);

        cell = tempTcSheet.getCell(6, 'b');
        cell.setCellFormula("COUNTA(B:B)-3");

        cell = tempTcSheet.getCell(6, 'c');
        cell.setCellFormula("COUNTIF(C:C,TRUE)");

        cell = tempTcSheet.getCell(6, 'd');
        cell.setCellFormula("COUNTIF(D:D,FALSE)");

        cell = tempTcSheet.getCell(6, 'e');
        cell.setCellFormula("COUNTIF(E:E,\"EX\") - 1");

        cell = tempTcSheet.getCell(6, 'f');
        cell.setCellFormula("COUNTIF(F:F,\"*TRUE*\")");

        cell = tempTcSheet.getCell(6, 'g');
        cell.setCellFormula("COUNTIF(G:G,\"*FALSE*\")");

        cell = tempTcSheet.getCell(6, 'h');
        cell.setCellFormula("COUNTIF(H:H,\"*TRUE\")");

        cell = tempTcSheet.getCell(6, 'i');
        cell.setCellFormula("COUNTIF(I:I,\"*FALSE*\")");

        cell = tempTcSheet.getCell(6, 'j');
        cell.setCellFormula("COUNTIF(J:J,\"*TRUE*\")");

        cell = tempTcSheet.getCell(6, 'k');
        cell.setCellFormula("COUNTIF(K:K,\"*FALSE*\")");

        cell = tempTcSheet.getCell(6, 'l');
        cell.setCellFormula("COUNTIF(L:L,\"*TRUE*\")");

        cell = tempTcSheet.getCell(6, 'm');
        cell.setCellFormula("COUNTIF(M:M,\"*FALSE*\")");

        cell = tempTcSheet.getCell(6, 'n');
        cell.setCellFormula("COUNTA(N:N)-2");

        cell = tempTcSheet.getCell(6, 'p');
        cell.setCellFormula("COUNTA(P:P)-2");

        cell = tempTcSheet.getCell(6, 'q');
        cell.setCellFormula("COUNTA(Q:Q)-2");

        cell = tempTcSheet.getCell(6, 'r');
        cell.setCellFormula("COUNTA(R:R)-2");

        cell = tempTcSheet.getCell(6, 's');
        cell.setCellFormula("COUNTA(S:S)-2");

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
        sheet.setColumnWidth('b', Constant.WAVSEP_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('p', Constant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('q', Constant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('r', Constant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('s', Constant.BENCHMARK_TEST_CASE_CELL_WIDTH);

        sheet.setColumnWidth('o', 3);
    }

    /**
     * initialize vulnerability sheet style
     *
     * @param sheet sheet to column width
     */
    private void initVulnerabilitySheetStyle(TcSheet sheet) {
        for (char i = 'a'; i <= 'b'; i++) {
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i), cellStyle.DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        }

        for (char i = 'c'; i <= 'm'; i++) {
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i), cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (char i = 'p'; i <= 's'; i++) {
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i), cellStyle.DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        }

        sheet.setSameCellStyle(3, 5, 'a', 'n', cellStyle.BOLD_CENTER_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(3, 5, 'p', 's', cellStyle.BOLD_CENTER_MIDDLE_RIGHT_LEFT);

        sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex("o"), cellStyle.DEFAULT_THICK_RIGHT_LEFT_BG_GRAY);

        sheet.setSameCellStyle(1, 2, 'a', 'n', cellStyle.BOLD_DEFAULT_THICK_TOP_BOTTOM);
        sheet.setSameCellStyle(1, 2, 'p', 's', cellStyle.BOLD_DEFAULT_THICK_TOP_BOTTOM);

        sheet.setSameCellStyle(6, 6, 'a', 'n', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(6, 6, 'p', 's', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);
    }


    /**
     * init vulnerability test cases
     *
     * @param tcSheet tcSheet to write
     */
    private void initVulnerabilityTC(TcSheet tcSheet, String vulnerabilityShortName) {
        List<String> tpTcList = FileUtil.readFile(Constant.TC_WAVSEP_PATH + vulnerabilityShortName + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION);
        int tpRowNum = TcUtil.writeDownListInSheet(tpTcList, tcSheet, 7, 'b');
        tcSheet.setSameCellStyle(7, tpRowNum - 1, 'b', 'b', cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        tcSheet.setSameValueMultiCells(7, tpRowNum - 1, 'c', 'c', true);
        tcSheet.setTpTcEndRowNum(tpRowNum - 1);

        List<String> fpTcList = FileUtil.readFile(Constant.TC_WAVSEP_PATH + vulnerabilityShortName + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION);
        int fpRowNum = TcUtil.writeDownListInSheet(fpTcList, tcSheet, tpRowNum, 'b');
        tcSheet.setSameCellStyle(tpRowNum, fpRowNum - 1, 'b', 'b', cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        tcSheet.setSameValueMultiCells(tpRowNum, fpRowNum - 1, 'd', 'd', false);
        tcSheet.setFpTcEndRowNum(fpRowNum - 1);

        List<String> exTcList = FileUtil.readFile(Constant.TC_WAVSEP_PATH + vulnerabilityShortName + Constant.EXPERIMENTAL_POSTFIX + Constant.TC_FILE_EXTENSION);
        int exRowNum = TcUtil.writeDownListInSheet(exTcList, tcSheet, fpRowNum, 'b');
        tcSheet.setSameCellStyle(fpRowNum, exRowNum - 1, 'b', 'b', cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        tcSheet.setSameValueMultiCells(fpRowNum, exRowNum - 1, 'e', 'e', "EX");
        tcSheet.setFpTcEndRowNum(exRowNum - 1);

        for (int i = 7; i < fpRowNum; i++) {
            String formula = "IF(EXACT(G" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            Cell cell = tcSheet.getCell(i, 'f');
            cell.setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = 7; i < fpRowNum; i++) {
            String formula = "IF(COUNTIF($P:$P,B" + i + ") > 0, \"FALSE\", \"\")";
            Cell cell = tcSheet.getCell(i, 'g');
            tcSheet.getCell(i, 'g').setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = 7; i < tpRowNum; i++) {
            String formula = "IF(EXACT(I" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            Cell cell = tcSheet.getCell(i, 'h');
            tcSheet.getCell(i, 'h').setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = 7; i < tpRowNum; i++) {
            String formula = "IF(COUNTIF($Q:$Q,B" + i + ") > 0, \"FALSE\", \"\")";
            Cell cell = tcSheet.getCell(i, 'i');
            tcSheet.getCell(i, 'i').setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = tpRowNum; i < fpRowNum; i++) {
            String formula = "IF(EXACT(K" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            Cell cell = tcSheet.getCell(i, 'j');
            tcSheet.getCell(i, 'j').setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = tpRowNum; i < fpRowNum; i++) {
            String formula = "IF(COUNTIF($R:$R,B" + i + ") > 0, \"FALSE\", \"\")";
            Cell cell = tcSheet.getCell(i, 'k');
            tcSheet.getCell(i, 'k').setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = fpRowNum; i < exRowNum; i++) {
            String formula = "IF(EXACT(M" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            Cell cell = tcSheet.getCell(i, 'l');
            tcSheet.getCell(i, 'l').setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = tpRowNum; i < fpRowNum; i++) {
            String formula = "IF(COUNTIF($S:$S,B" + i + ") > 0, \"FALSE\", \"\")";
            Cell cell = tcSheet.getCell(i, 'm');
            tcSheet.getCell(i, 'm').setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }
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

        cell = tempTcSheet.getCell(1, 'b');
        cell.setCellValue("TP");

        cell = tempTcSheet.getCell(1, 'c');
        cell.setCellValue("FN");

        cell = tempTcSheet.getCell(1, 'd');
        cell.setCellValue("TN");

        cell = tempTcSheet.getCell(1, 'e');
        cell.setCellValue("FP");

        cell = tempTcSheet.getCell(1, 'f');
        cell.setCellValue("Total");

        cell = tempTcSheet.getCell(1, 'g');
        cell.setCellValue("TPR");

        cell = tempTcSheet.getCell(1, 'h');
        cell.setCellValue("FPR");

        cell = tempTcSheet.getCell(1, 'i');
        cell.setCellValue("Score");

        cell = tempTcSheet.getCell(2, 'a');
        cell.setCellValue("Vulnerabilities + False Positive");

        int rowNum = 2;

        for (Constant.WavsepSheets vulnerability : Constant.WavsepSheets.values()) {
            if (!vulnerability.toString().equals("total")) {
                rowNum++;
                cell = tempTcSheet.getCell(rowNum, 'a');
                cell.setCellValue(vulnerability.getSheetName());
            }
        }

        cell = tempTcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("Experimental Test Cases (inspired/imported from ZAP-WAVE)");

        cell = tempTcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("additional RXSS(anticsrf tokens, secret input vectors, tag signatures etc.)");

        cell = tempTcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("addtional SQLi(INSERT)");

        cell = tempTcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("Totals");

        cell = tempTcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("Overall Results");
    }

    private void initTotalColumnWidth(TcSheet sheet) {
        sheet.getSheet().setDefaultColumnWidth(Constant.DEFAULT_COLUMN_WIDTH);
        sheet.setColumnWidth('a', Constant.WAVSEP_TEST_CASE_CELL_WIDTH);
    }

    private void initTotalSheetStyle(TcSheet sheet) {
        sheet.setSameCellStyle(1, 1, 'a', 'i', CellStylesUtil.getBoldCenterStyle(this.tcWorkbook.getWorkbook()));
    }

    //todo define default XSSFCEllStyle

    /**
     * write current state in excel file - ZAP_WAVSEP_Results_[time].xlsx
     */
    public void writeWavsep() {
        if (tcWorkbook != null) {
            tcWorkbook.writeWorkbook(Constant.WAVSEP_NAME);
        }
    }
}
