package org.poi.BenchmarkPOI;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.Constant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.TcUtil;
import org.apache.poi.ss.usermodel.*;
import org.poi.POIConstant;
import org.poi.Util.CellStylesUtil;
import org.poi.Util.FileUtil;
import org.poi.WavsepPOI.WavsepParser;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * Created by Hwan on 2016-05-10.
 */
public class BenchmarkTemplate {
    private TcWorkbook tcWorkbook;
    private CellStylesUtil cellStyle;

    private String lastWriteFilePath;

    public BenchmarkTemplate() {
        this.tcWorkbook = new TcWorkbook();
        this.cellStyle = new CellStylesUtil(this.tcWorkbook.getWorkbook());
        init();
    }

    public BenchmarkTemplate(XSSFWorkbook workbook){
        this.tcWorkbook = new TcWorkbook(workbook);
        this.cellStyle = new CellStylesUtil(this.tcWorkbook.getWorkbook());
    }

    /**
     * Set Benchmark workbook templates.
     */
    private void init() {
        for (Constant.BenchmarkSheets benchmarkSheet : Constant.BenchmarkSheets.values()) {
            if (!benchmarkSheet.getSheetName().equals(Constant.BenchmarkSheets.total.getSheetName())) {
                try {
                    initVulnerabilitySheet(this.tcWorkbook.getTcSheet(benchmarkSheet.getSheetName()));
                    initVulnerabilityTC(this.tcWorkbook.getTcSheet(benchmarkSheet.getSheetName()), benchmarkSheet.toString());
                } catch (IOException e){
                    System.out.println("BenchmarkTemplate : can't write vulnerability Tc since occur a IOException - " + benchmarkSheet.getSheetName());
                    e.printStackTrace();
                }
            } else {
                initTotalSheet(this.tcWorkbook.getTcSheet(benchmarkSheet.getSheetName()));
            }
        }

    }

    /**
     * generate vulnerability sheet template
     *
     * @param tcSheet vulnerability sheet
     */
    private void initVulnerabilitySheet(TcSheet tcSheet) {
        initVulnerabilityColumnWidth(tcSheet);

        CreationHelper createHelper = tcWorkbook.getWorkbook().getCreationHelper();

        Cell cell = tcSheet.getCell(1, 'A');
        cell.setCellValue("검사날짜");

        cell = tcSheet.getCell(1, 'B');
        // Actually this is not 검사날짜.
        cell.setCellValue(tcWorkbook.getWorkbookCreatedTime());

        cell = tcSheet.getCell(2, 'a');
        cell.setCellValue(tcSheet.getSheet().getSheetName());

        cell = tcSheet.mergeCellRegion(3, 5, 'b', 'b');
        cell.setCellValue("Test Cases");

        cell = tcSheet.mergeCellRegion(3, 4, 'c', 'd');
        cell.setCellValue("Real Vulnerability");

        cell = tcSheet.mergeCellRegion(3, 4, 'e', 'f');
        cell.setCellValue("URL Crawl");

        cell = tcSheet.mergeCellRegion(3, 3, 'g', 'j');
        cell.setCellValue("Detected");

        cell = tcSheet.mergeCellRegion(4, 4, 'g', 'h');
        cell.setCellValue("TRUE Vulnerability");

        cell = tcSheet.mergeCellRegion(4, 4, 'i', 'j');
        cell.setCellValue("FALSE Vulnerability");

        cell = tcSheet.mergeCellRegion(3, 5, 'k', 'k');
        cell.setCellValue("Description");

        cell = tcSheet.mergeCellRegion(3, 5, 'm', 'm');
        cell.setCellValue("Failed Crawling");

        cell = tcSheet.mergeCellRegion(3, 5, 'n', 'n');
        cell.setCellValue("Failed TP");

        cell = tcSheet.mergeCellRegion(3, 5, 'o', 'o');
        cell.setCellValue("Failed FP");

        cell = tcSheet.getCell(6, 'a');
        cell.setCellValue("합계");

        cell = tcSheet.getCell(5, 'c');
        cell.setCellValue(true);

        cell = tcSheet.getCell(5, 'd');
        cell.setCellValue(false);

        cell = tcSheet.getCell(5, 'e');
        cell.setCellValue(true);

        cell = tcSheet.getCell(5, 'f');
        cell.setCellValue(false);

        cell = tcSheet.getCell(5, 'g');
        cell.setCellValue(true);

        cell = tcSheet.getCell(5, 'h');
        cell.setCellValue(false);

        cell = tcSheet.getCell(5, 'i');
        cell.setCellValue(true);

        cell = tcSheet.getCell(5, 'j');
        cell.setCellValue(false);

        cell = tcSheet.getCell(6, 'b');
        cell.setCellFormula("COUNTA(B:B)-3");

        cell = tcSheet.getCell(6, 'c');
        cell.setCellFormula("COUNTIF(C:C,TRUE)-1");

        cell = tcSheet.getCell(6, 'd');
        cell.setCellFormula("COUNTIF(D:D,FALSE)-1");

        cell = tcSheet.getCell(6, 'e');
        cell.setCellFormula("COUNTIF(E:E,\"*TRUE*\")");

        cell = tcSheet.getCell(6, 'f');
        cell.setCellFormula("COUNTIF(F:F,\"*FALSE*\")");

        cell = tcSheet.getCell(6, 'g');
        cell.setCellFormula("COUNTIF(G:G,\"*TRUE*\")-1");

        cell = tcSheet.getCell(6, 'h');
        cell.setCellFormula("COUNTIF(H:H,\"*FALSE*\")");

        cell = tcSheet.getCell(6, 'i');
        cell.setCellFormula("COUNTIF(I:I,\"*TRUE*\")");

        cell = tcSheet.getCell(6, 'j');
        cell.setCellFormula("COUNTIF(J:J,\"*FALSE*\")");

        cell = tcSheet.getCell(6, 'k');
        cell.setCellFormula("COUNTA(K:K)-2");

        cell = tcSheet.getCell(6, 'm');
        cell.setCellFormula("COUNTA(M:M)-2");

        cell = tcSheet.getCell(6, 'n');
        cell.setCellFormula("COUNTA(N:N)-2");

        cell = tcSheet.getCell(6, 'o');
        cell.setCellFormula("COUNTA(O:O)-2");

        initVulnerabilitySheetStyle(tcSheet);
        cell = tcSheet.getCell(1, 'B');
        cell.getCellStyle().setDataFormat(createHelper.createDataFormat().getFormat("yyyy-MM-dd HH:mm"));
    }

    /**
     * initialize vulnerability column width
     *
     * @param sheet sheet to set column width
     */
    private void initVulnerabilityColumnWidth(TcSheet sheet) {
        for(char i = 'c'; i <= 'j'; i++){
            sheet.setColumnWidth(i, POIConstant.DEFAULT_COLUMN_WIDTH);
        }
        sheet.setColumnWidth('a',POIConstant.DEFAULT_COLUMN_WIDTH);
        sheet.setColumnWidth('b', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('m', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('n', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        sheet.setColumnWidth('o', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);

        sheet.setColumnWidth('l', 3);
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

        for (char i = 'c'; i <= 'j'; i++) {
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i), cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (char i = 'm'; i <= 'o'; i++) {
            sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex(i), cellStyle.DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        }

        sheet.setSameCellStyle(3, 5, 'a', 'k', cellStyle.BOLD_CENTER_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(3, 5, 'm', 'o', cellStyle.BOLD_CENTER_MIDDLE_RIGHT_LEFT);

        sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex("l"), cellStyle.DEFAULT_THICK_RIGHT_LEFT_BG_GRAY);

        sheet.setSameCellStyle(1, 2, 'a', 'k', cellStyle.BOLD_DEFAULT_THICK_TOP_BOTTOM);
        sheet.setSameCellStyle(1, 2, 'm', 'o', cellStyle.BOLD_DEFAULT_THICK_TOP_BOTTOM);

        sheet.setSameCellStyle(6, 6, 'a', 'k', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);
        sheet.setSameCellStyle(6, 6, 'm', 'o', cellStyle.BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT);
    }

    /**
     * init vulnerability test cases
     *
     * @param tcSheet tcSheet to write
     */
    private void initVulnerabilityTC(TcSheet tcSheet, String vulnerabilityShortName) throws IOException {
        List<String> tpTcList = FileUtil.readResourceFile(BenchmarkTemplate.class, BenchmarkParser.BENCHMARK_TC_PATH + vulnerabilityShortName + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION);
        int tpRowNum = TcUtil.writeDownListInSheet(tpTcList, tcSheet, 7, 'b');
        tcSheet.setSameValueMultiCells(7, tpRowNum - 1, 'c', 'c', true);
        tcSheet.setSameCellStyle(7, tpRowNum - 1, 'c', 'c', cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        tcSheet.setTpTcEndRowNum(tpRowNum - 1);

        List<String> fpTcList = FileUtil.readResourceFile(BenchmarkTemplate.class, BenchmarkParser.BENCHMARK_TC_PATH + vulnerabilityShortName + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION);
        int fpRowNum = TcUtil.writeDownListInSheet(fpTcList, tcSheet, tpRowNum, 'b');
        tcSheet.setSameValueMultiCells(tpRowNum, fpRowNum - 1, 'd', 'd', false);
        tcSheet.setSameCellStyle(tpRowNum, fpRowNum - 1, 'd', 'd', cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        tcSheet.setFpTcEndRowNum(fpRowNum - 1);


        for (int i = 7; i < fpRowNum; i++) {
            String formula = "IF(EXACT(F" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            Cell cell = tcSheet.getCell(i, 'e');
            cell.setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = 7; i < fpRowNum; i++) {
            String formula = "IF(COUNTIF($M:$M,B" + i + ") > 0, \"FALSE\", \"\")";
            Cell cell = tcSheet.getCell(i, 'f');
            cell.setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = 7; i < tpRowNum; i++) {
            String formula = "IF(EXACT(H" + i + ",\"FALSE\"), \"\", \"TRUE\")";
            Cell cell = tcSheet.getCell(i, 'g');
            cell.setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = 7; i < tpRowNum; i++) {
            String formula = "IF(COUNTIF($N:$N,B" + i + ") > 0, \"FALSE\", \"\")";
            Cell cell = tcSheet.getCell(i, 'h');
            cell.setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = tpRowNum; i < fpRowNum; i++) {
            String formula = "IF(COUNTIF($O:$O,B" + i + ") > 0, \"TRUE\", \"\")";
            Cell cell = tcSheet.getCell(i, 'i');
            cell.setCellFormula(formula);
            cell.setCellStyle(cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        for (int i = tpRowNum; i < fpRowNum; i++) {
            String formula = "IF(EXACT(I" + i + ",\"TRUE\"), \"\", \"FALSE\")";
            Cell cell = tcSheet.getCell(i, 'j');
            cell.setCellFormula(formula);
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

        cell = tcSheet.getCell(1, 'B');
        cell.setCellValue("CWE #");

        cell = tcSheet.getCell(1, 'c');
        cell.setCellValue("TP");

        cell = tcSheet.getCell(1, 'd');
        cell.setCellValue("FN");

        cell = tcSheet.getCell(1, 'e');
        cell.setCellValue("TN");

        cell = tcSheet.getCell(1, 'f');
        cell.setCellValue("FP");

        cell = tcSheet.getCell(1, 'g');
        cell.setCellValue("Total");

        cell = tcSheet.getCell(1, 'h');
        cell.setCellValue("TPR");

        cell = tcSheet.getCell(1, 'i');
        cell.setCellValue("FPR");

        cell = tcSheet.getCell(1, 'j');
        cell.setCellValue("Score");

        int rowNum = 1;

        for (Constant.BenchmarkSheets vulnerability : Constant.BenchmarkSheets.values()) {
            if (!vulnerability.getSheetName().equals("Total")) {
                rowNum++;
                cell = tcSheet.getCell(rowNum, 'a');
                cell.setCellValue(vulnerability.getSheetName());
                cell = tcSheet.getCell(rowNum, 'b');
                cell.setCellValue(vulnerability.getCwe());

                cell = tcSheet.getCell(rowNum, 'c');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!G6");
                cell = tcSheet.getCell(rowNum, 'd');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!h6");
                cell = tcSheet.getCell(rowNum, 'e');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!i6");
                cell = tcSheet.getCell(rowNum, 'f');
                cell.setCellFormula("\'" + vulnerability.getSheetName() + "\'!j6");

                cell = tcSheet.getCell(rowNum, 'g');
                cell.setCellFormula("sum(C" + rowNum + ":F" + rowNum + ")");
                cell = tcSheet.getCell(rowNum, 'h');
                cell.setCellFormula("C" + rowNum + "/(C" + rowNum + "+D" + rowNum + ")");
                cell = tcSheet.getCell(rowNum, 'i');
                cell.setCellFormula("F" + rowNum + "/(F" + rowNum + "+E" + rowNum + ")");
                cell = tcSheet.getCell(rowNum, 'j');
                cell.setCellFormula("H" + rowNum + "-I" + rowNum);
            }
        }

        cell = tcSheet.getCell(13, 'c');
        cell.setCellFormula("sum(C2:C12)");
        cell = tcSheet.getCell(13, 'd');
        cell.setCellFormula("sum(D2:D12)");
        cell = tcSheet.getCell(13, 'e');
        cell.setCellFormula("sum(E2:E12)");
        cell = tcSheet.getCell(13, 'f');
        cell.setCellFormula("sum(F2:F12)");
        cell = tcSheet.getCell(13, 'G');
        cell.setCellFormula("sum(G2:G12)");

        cell = tcSheet.getCell(14, 'h');
        cell.setCellFormula("AVERAGE(H2:H12)");
        cell = tcSheet.getCell(14, 'i');
        cell.setCellFormula("AVERAGE(I2:I12)");
        cell = tcSheet.getCell(14, 'j');
        cell.setCellFormula("AVERAGE(J2:J12)");

        cell = tcSheet.getCell(13, 'a');
        cell.setCellValue("Totals");

        cell = tcSheet.getCell(14, 'a');
        cell.setCellValue("Overall Results");
    }

    private void initTotalColumnWidth(TcSheet sheet) {
        sheet.setColumnWidth('a', 30);
        sheet.getSheet().setDefaultColumnWidth(POIConstant.DEFAULT_COLUMN_WIDTH);
    }

    /**
     * set total sheet style
     *
     * @param sheet total sheet
     */
    private void initTotalSheetStyle(TcSheet sheet) {
        CellStyle boldCenterThickBottomRight = cellStyle.getSimpleCellStyle(true, true, 0, 2, 0, 2);
        CellStyle boldCenterThickBottom = cellStyle.getSimpleCellStyle(true, true, 0, 2, 0, 0);
        CellStyle thickBottomRight = cellStyle.getSimpleCellStyle(false, false, 0, 2, 0, 2);
        CellStyle thinTopBotThickRight = cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 2);
        CellStyle thinTopBot = cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 0);
        CellStyle thickBottom = cellStyle.getSimpleCellStyle(false, false, 0, 2, 0, 0);

        Cell cell = sheet.getCell(1, 'a');
        cell.setCellStyle(boldCenterThickBottomRight);
        cell = sheet.getCell(1, 'g');
        cell.setCellStyle(boldCenterThickBottomRight);
        cell = sheet.getCell(1, 'j');
        cell.setCellStyle(boldCenterThickBottomRight);

        sheet.setSameCellStyle(2, 11, 'a', 'a', thinTopBotThickRight);
        sheet.setSameCellStyle(2, 11, 'g', 'g', thinTopBotThickRight);

        cell = sheet.getCell(13, 'a');
        cell.setCellStyle(thinTopBotThickRight);
        cell = sheet.getCell(13, 'g');
        cell.setCellStyle(thinTopBotThickRight);

        cell = sheet.getCell(12, 'a');
        cell.setCellStyle(thickBottomRight);
        cell = sheet.getCell(12, 'g');
        cell.setCellStyle(thickBottomRight);
        cell = sheet.getCell(14, 'a');
        cell.setCellStyle(thickBottomRight);
        cell = sheet.getCell(14, 'g');
        cell.setCellStyle(thickBottomRight);

        sheet.setSameCellStyle(1, 1, 'b', 'f', boldCenterThickBottom);
        sheet.setSameCellStyle(1, 1, 'h', 'i', boldCenterThickBottom);

        sheet.setSameCellStyle(12, 12, 'b', 'f', thickBottom);
        sheet.setSameCellStyle(14, 14, 'b', 'f', thickBottom);

        sheet.setSameCellStyle(2, 11, 'b', 'f', thinTopBot);
        sheet.setSameCellStyle(13, 13, 'b', 'f', thinTopBot);

        /** set percentage cell style */
        CellStyle pThinTopBotThickRight = cellStyle.getSimpleCellStyle(false, false);
        pThinTopBotThickRight.cloneStyleFrom(thinTopBotThickRight);
        pThinTopBotThickRight.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        CellStyle pThickBottomRight = cellStyle.getSimpleCellStyle(false, false);
        pThickBottomRight.cloneStyleFrom(thickBottomRight);
        pThickBottomRight.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        CellStyle pThickBottom = cellStyle.getSimpleCellStyle(false, false);
        pThickBottom.cloneStyleFrom(thickBottom);
        pThickBottom.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        CellStyle pThinTopBot = cellStyle.getSimpleCellStyle(false, false);
        pThinTopBot.cloneStyleFrom(thinTopBot);
        pThinTopBot.setDataFormat(this.tcWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        cell = sheet.getCell(13, 'j');
        cell.setCellStyle(pThinTopBotThickRight);
        sheet.setSameCellStyle(2, 11, 'j', 'j', pThinTopBotThickRight);

        cell = sheet.getCell(12, 'j');
        cell.setCellStyle(pThickBottomRight);
        cell = sheet.getCell(14, 'j');
        cell.setCellStyle(pThickBottomRight);

        sheet.setSameCellStyle(12, 12, 'h', 'i', pThickBottom);
        sheet.setSameCellStyle(14, 14, 'h', 'i', pThickBottom);

        sheet.setSameCellStyle(2, 11, 'h', 'i', pThinTopBot);
        sheet.setSameCellStyle(13, 13, 'h', 'i', pThinTopBot);
    }

    /**
     * write current state in excel file - ZAP_OWASPBenchmark_Results_[time].xlsx
     */
    public void writeBenchmark() {
        if (tcWorkbook != null) {
            lastWriteFilePath = tcWorkbook.writeWorkbook(POIConstant.BENCHMARK_NAME);
        }
    }

    /**
     * write current state in cxcel file
     *
     * @param outputDir output directory path
     * @param fileName file name
     * @return file absoluete path
     */
    public String writeBenchmark(String outputDir, String fileName){
        if(tcWorkbook != null){
            lastWriteFilePath = tcWorkbook.writeWorkbook(outputDir, fileName);
            return lastWriteFilePath;
        }
        return null;
    }

//    public String getLastWriteFilePath() {
//        return lastWriteFilePath;
//    }
//
//    public String appendBenchmarkFailedList(String crawledPath, String scanPath){
//        if(lastWriteFilePath == null){
//            System.err.println("BenchmarkTemplate : Append failed list error. First, you must write benchmark to append.");
//            return null;
//        }
//
//    }

    /**
     * write down failed lists
     *
     * @param resTcDirPath result test cases directory path
     */
    public void writeFailedList(String resTcDirPath) {
        File resDir = new File(resTcDirPath);

        if (resDir.exists() && resDir.isDirectory()) {
            for (File tcFile : resDir.listFiles()) {
                String fileName = tcFile.getName();
                String vulnerability;
                int index;
                boolean crawlFlag = false;

                if ((index = fileName.indexOf(BenchmarkParser.BENCHMARK_TEST_PREFIX+ BenchmarkParser.BENCHMARK_TEST_CRAWLED) ) > 0) {
                    vulnerability = fileName.substring(index + (BenchmarkParser.BENCHMARK_TEST_PREFIX+ BenchmarkParser.BENCHMARK_TEST_CRAWLED).length(), fileName.lastIndexOf("."));
                    crawlFlag = true;
                } else if ((index = fileName.indexOf(BenchmarkParser.BENCHMARK_TEST_PREFIX)) > 0){
                    vulnerability = fileName.substring(index + BenchmarkParser.BENCHMARK_TEST_PREFIX.length(), fileName.lastIndexOf("."));
                } else {
                    continue;
                }

                for(Constant.BenchmarkSheets benchmark : Constant.BenchmarkSheets.values()){
                    if(benchmark.toString().equals(vulnerability)){
                        TcSheet tcSheet = this.tcWorkbook.getTcSheet(benchmark.getSheetName());
                        try {
                            if (crawlFlag) {
                                TcUtil.writeDownListInSheet(BenchmarkParser.getFailedCrawlingList(tcFile, vulnerability), tcSheet, 7, 'm');
                            } else {
                                TcUtil.writeDownListInSheet(BenchmarkParser.getFailedTPList(tcFile.getAbsolutePath(), vulnerability), tcSheet, 7, 'n');
                                TcUtil.writeDownListInSheet(BenchmarkParser.getFailedFPList(tcFile.getAbsolutePath(), vulnerability), tcSheet, 7, 'o');
                            }
                        } catch (IOException e) {
                            System.err.println("BenchmarkTemplate : Occur a IOException - " + fileName);
                            e.printStackTrace();
                        }
                        break;
                    }
                }
            }
        }
    }
}