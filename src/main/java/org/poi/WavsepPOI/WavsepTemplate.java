package org.poi.WavsepPOI;

import org.poi.Constant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.TcUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.poi.POIConstant;
import org.poi.Util.CellStylesUtil;
import org.poi.Util.FileUtil;

import java.io.File;
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
        for(char i = 'c'; i <= 'n'; i++){
            sheet.setColumnWidth(i, POIConstant.DEFAULT_COLUMN_WIDTH);
        }
        sheet.setColumnWidth('a', POIConstant.DEFAULT_COLUMN_WIDTH);
        sheet.setColumnWidth('b', POIConstant.WAVSEP_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('p', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('q', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('r', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('s', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);

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
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i), cellStyle.BOLD_CENTER_MIDDLE_RIGHT_LEFT);
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
     * @param tcSheet total sheet
     */
    private void initTotalSheet(TcSheet tcSheet) {

        initTotalColumnWidth(tcSheet);
        initTotalSheetStyle(tcSheet);

        Cell cell = tcSheet.getCell(1, 'A');
        cell.setCellValue("Category");

        cell = tcSheet.getCell(1, 'b');
        cell.setCellValue("TP");

        cell = tcSheet.getCell(1, 'c');
        cell.setCellValue("FN");

        cell = tcSheet.getCell(1, 'd');
        cell.setCellValue("TN");

        cell = tcSheet.getCell(1, 'e');
        cell.setCellValue("FP");

        cell = tcSheet.getCell(1, 'f');
        cell.setCellValue("Total");

        cell = tcSheet.getCell(1, 'g');
        cell.setCellValue("TPR");

        cell = tcSheet.getCell(1, 'h');
        cell.setCellValue("FPR");

        cell = tcSheet.getCell(1, 'i');
        cell.setCellValue("Score");

        cell = tcSheet.getCell(2, 'a');
        cell.setCellValue("Vulnerabilities + False Positive");

        int rowNum = 2;

        for (Constant.WavsepSheets vulnerability : Constant.WavsepSheets.values()) {
            if (!vulnerability.toString().equals("total")) {
                rowNum++;
                cell = tcSheet.getCell(rowNum, 'a');
                cell.setCellValue(vulnerability.getSheetName());

                cell = tcSheet.getCell(rowNum, 'b');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!H6");
                cell = tcSheet.getCell(rowNum, 'c');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!I6");
                cell = tcSheet.getCell(rowNum, 'd');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!J6");
                cell = tcSheet.getCell(rowNum, 'e');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!K6");

                cell = tcSheet.getCell(rowNum, 'f');
                cell.setCellFormula("SUM(B" + rowNum + ":E" + rowNum + ")");
                cell = tcSheet.getCell(rowNum, 'g');
                cell.setCellFormula("B" + rowNum + "/(B" + rowNum + "+C" + rowNum + ")");
                cell = tcSheet.getCell(rowNum, 'h');
                cell.setCellFormula("E" + rowNum + "/(E" + rowNum + "+D" + rowNum + ")");
                cell = tcSheet.getCell(rowNum, 'i');
                cell.setCellFormula("G" + rowNum + "-H" + rowNum);

            }

            if (vulnerability.toString().equals("rxss")) {
                cell = tcSheet.getCell(9, 'b');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!L6");
                cell = tcSheet.getCell(9, 'c');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!M6");

                cell = tcSheet.getCell(9, 'f');
                cell.setCellFormula("B9+C9");
                cell = tcSheet.getCell(9, 'g');
                cell.setCellFormula("B9/(B9+C9)");
                cell = tcSheet.getCell(9, 'i');
                cell.setCellFormula("G9");
            } else if (vulnerability.toString().equals("sqli")) {
                cell = tcSheet.getCell(10, 'b');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!L6");
                cell = tcSheet.getCell(10, 'c');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!M6");

                cell = tcSheet.getCell(10, 'f');
                cell.setCellFormula("B10+C10");
                cell = tcSheet.getCell(10, 'g');
                cell.setCellFormula("B10/(B10+C10)");
                cell = tcSheet.getCell(10, 'i');
                cell.setCellFormula("G10");
            }
        }

        cell = tcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("Experimental Test Cases (inspired/imported from ZAP-WAVE)");

        cell = tcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("additional RXSS(anticsrf tokens, secret input vectors, tag signatures etc.)");

        cell = tcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("addtional SQLi(INSERT)");

        cell = tcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("Totals");

        cell = tcSheet.getCell(++rowNum, 'a');
        cell.setCellValue("Overall Results");

        cell = tcSheet.getCell(11, 'b');
        cell.setCellFormula("SUM(B3:B7, B9:B10)");
        cell = tcSheet.getCell(11, 'c');
        cell.setCellFormula("SUM(C3:C7, C9:C10)");
        cell = tcSheet.getCell(11, 'd');
        cell.setCellFormula("SUM(D3:D7, D9:D10)");
        cell = tcSheet.getCell(11, 'e');
        cell.setCellFormula("SUM(E3:E7, E9:E10)");
        cell = tcSheet.getCell(11, 'f');
        cell.setCellFormula("SUM(F3:F7, F9:F10)");

        cell = tcSheet.getCell(12, 'g');
        cell.setCellFormula("AVERAGE(G3:G7,G9:G10)");
        cell = tcSheet.getCell(12, 'h');
        cell.setCellFormula("AVERAGE(H3:H7)");
        cell = tcSheet.getCell(12, 'i');
        cell.setCellFormula("AVERAGE(I3:I7,I9:I10)");
    }

    private void initTotalColumnWidth(TcSheet sheet) {
        sheet.getSheet().setDefaultColumnWidth(POIConstant.DEFAULT_COLUMN_WIDTH);
        sheet.setColumnWidth('a', POIConstant.WAVSEP_TEST_CASE_CELL_WIDTH);
    }

    private void initTotalSheetStyle(TcSheet sheet) {
        CellStyle boldCenterThickRight = cellStyle.getSimpleCellStyle(true, true, 0, 0, 0, 2);
        CellStyle thickRight = cellStyle.getSimpleCellStyle(false, false, 0, 0, 0, 2);
        CellStyle thinUpThickBotRight = cellStyle.getSimpleCellStyle(false, false, 1, 2, 0, 2);
        CellStyle thickUpBotRight = cellStyle.getSimpleCellStyle(false, false, 2, 2, 0, 2);
        CellStyle thinTopBot = cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 0);
        CellStyle thinTopThickBot = cellStyle.getSimpleCellStyle(false, false, 1, 2, 0, 0);
        CellStyle thinTopBotThickRight = cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 2);

        Cell cell = sheet.getCell(1, 'a');
        cell.setCellStyle(boldCenterThickRight);
        cell = sheet.getCell(1, 'f');
        cell.setCellStyle(boldCenterThickRight);
        cell = sheet.getCell(1, 'i');
        cell.setCellStyle(boldCenterThickRight);

        cell = sheet.getCell(3, 'a');
        cell.setCellStyle(thickRight);
        cell = sheet.getCell(7, 'a');
        cell.setCellStyle(thickRight);
        cell = sheet.getCell(9, 'a');
        cell.setCellStyle(thickRight);
        cell = sheet.getCell(11, 'a');
        cell.setCellStyle(thickRight);

        cell = sheet.getCell(3, 'f');
        cell.setCellStyle(thickRight);
        cell = sheet.getCell(7, 'f');
        cell.setCellStyle(thickRight);
        cell = sheet.getCell(9, 'f');
        cell.setCellStyle(thickRight);
        cell = sheet.getCell(11, 'f');
        cell.setCellStyle(thickRight);

        sheet.setSameCellStyle(2, 2, 'a', 'h', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM);
        cell = sheet.getCell(2, 'i');
        cell.setCellStyle(thickUpBotRight);
        sheet.setSameCellStyle(8, 8, 'a', 'h', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM);
        cell = sheet.getCell(8, 'i');
        cell.setCellStyle(thickUpBotRight);
        sheet.setSameCellStyle(1, 1, 'b', 'e', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM);
        sheet.setSameCellStyle(1, 1, 'g', 'h', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM);

        cell = sheet.getCell(10, 'a');
        cell.setCellStyle(thinUpThickBotRight);
        cell = sheet.getCell(12, 'a');
        cell.setCellStyle(thinUpThickBotRight);
        cell = sheet.getCell(10, 'f');
        cell.setCellStyle(thinUpThickBotRight);
        cell = sheet.getCell(12, 'f');
        cell.setCellStyle(thinUpThickBotRight);

        sheet.setSameCellStyle(4, 6, 'b', 'e', thinTopBot);

        sheet.setSameCellStyle(10, 10, 'b', 'e', thinTopThickBot);
        sheet.setSameCellStyle(12, 12, 'b', 'e', thinTopThickBot);

        sheet.setSameCellStyle(4, 6, 'a', 'a', thinTopBotThickRight);
        sheet.setSameCellStyle(4, 6, 'f', 'f', thinTopBotThickRight);

        /** set percentage cell style */
        CellStyle pThickRight = this.tcWorkbook.getWorkbook().createCellStyle();
        pThickRight.cloneStyleFrom(thickRight);
        pThickRight.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        CellStyle pThinTopBotThickRight = this.tcWorkbook.getWorkbook().createCellStyle();
        pThinTopBotThickRight.cloneStyleFrom(thinTopBotThickRight);
        pThinTopBotThickRight.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        CellStyle pThinUpThickBotRight = this.tcWorkbook.getWorkbook().createCellStyle();
        pThinUpThickBotRight.cloneStyleFrom(thinUpThickBotRight);
        pThinUpThickBotRight.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        CellStyle pThinTopBot = this.tcWorkbook.getWorkbook().createCellStyle();
        pThinTopBot.cloneStyleFrom(thinTopBot);
        pThinTopBot.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        CellStyle pThinTopThickBot = this.tcWorkbook.getWorkbook().createCellStyle();
        pThinTopThickBot.cloneStyleFrom(thinTopThickBot);
        pThinTopThickBot.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        CellStyle p = cellStyle.getSimpleCellStyle(false, false);
        p.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        cell = sheet.getCell(3, 'i');
        cell.setCellStyle(pThickRight);
        cell = sheet.getCell(7, 'i');
        cell.setCellStyle(pThickRight);
        cell = sheet.getCell(9, 'i');
        cell.setCellStyle(pThickRight);
        cell = sheet.getCell(11, 'i');
        cell.setCellStyle(pThickRight);

        sheet.setSameCellStyle(4, 6, 'i', 'i', pThinTopBotThickRight);

        cell = sheet.getCell(10, 'i');
        cell.setCellStyle(pThinUpThickBotRight);
        cell = sheet.getCell(12, 'i');
        cell.setCellStyle(pThinUpThickBotRight);

        sheet.setSameCellStyle(4, 6, 'g', 'h', pThinTopBot);

        sheet.setSameCellStyle(12, 12, 'g', 'h', pThinTopThickBot);
        sheet.setSameCellStyle(10, 10, 'g', 'h', pThinTopThickBot);

        cell = sheet.getCell(3, 'g');
        cell.setCellStyle(p);
        cell = sheet.getCell(7, 'g');
        cell.setCellStyle(p);
        cell = sheet.getCell(9, 'g');
        cell.setCellStyle(p);
        cell = sheet.getCell(3, 'h');
        cell.setCellStyle(p);
        cell = sheet.getCell(7, 'h');
        cell.setCellStyle(p);
    }

    //todo define default XSSFCEllStyle


    /**
     * write current state in excel file - ZAP_WAVSEP_Results_[time].xlsx
     *
     * @return file absolute path
     */
    public String writeWavsep() {
        if (tcWorkbook != null) {
            return tcWorkbook.writeWorkbook(POIConstant.WAVSEP_NAME);
        }
        return null;
    }

    /**
     * write current state in excel file
     *
     * @param outputDir output directory path
     * @param fileName file name
     * @return file absolute path
     */
    public String writeWavsep(String outputDir, String fileName){
        if(tcWorkbook != null){
            return tcWorkbook.writeWorkbook(outputDir, fileName);
        }
        return null;
    }

    /**
     * set Failed List - failed Crawling, failed TP, failed FP, failed EX
     *
     * @param resTcDirPath scanning results directory path
     */
    public void setFailedList(String resTcDirPath) {
        File resDir = new File(resTcDirPath);

        if (resDir.exists() && resDir.isDirectory()) {
            for (File tcFile : resDir.listFiles()) {
                String fileName = tcFile.getName();
                String vulnerability;
                int index;
                boolean crawlFlag = false;

                if ((index = fileName.indexOf(WavsepParser.WAVSEP_TEST_PREFIX + WavsepParser.WAVSEP_TEST_CRAWLED) ) > 0) {
                    vulnerability = fileName.substring(index);
                    crawlFlag = true;
                } else if ((index = fileName.indexOf(WavsepParser.WAVSEP_TEST_PREFIX)) > 0){
                    vulnerability = fileName.substring(index);
                } else {
                    continue;
                }

                for(Constant.WavsepSheets wavsep : Constant.WavsepSheets.values()){
                    if(wavsep.toString().equals(vulnerability)){
                        TcSheet tcSheet = this.tcWorkbook.getTcSheet(wavsep.getSheetName());
                        if(crawlFlag){
                            TcUtil.writeDownListInSheet(WavsepParser.getFailedCrawlingList(tcFile), tcSheet, 7, 'p');
                        } else {
                            TcUtil.writeDownListInSheet(WavsepParser.getFailedTcList(tcFile, vulnerability, Constant.TEST_SET_FAILED_TRUE_POSITIVE), tcSheet, 7, 'q');
                            TcUtil.writeDownListInSheet(WavsepParser.getFailedTcList(tcFile, vulnerability, Constant.TEST_SET_FAILED_FALSE_POSITIVE), tcSheet, 7, 'r');
                            TcUtil.writeDownListInSheet(WavsepParser.getFailedTcList(tcFile, vulnerability, Constant.TEST_SET_FAILED_EXPERIMENTAL), tcSheet, 7, 's');
                        }
                    }
                }
            }
        }
    }
}
