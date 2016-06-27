package org.poi.WavsepPOI;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.Constant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.CellStylesUtil;
import org.poi.Util.FileUtil;
import org.poi.Util.TcUtil;

import java.io.File;
import java.util.*;

/**
 * Created by Hwan on 2016-06-21.
 */
public class WavsepWriter {

    private String wavsepTemplateFilePath;
    private TcWorkbook wavsepWorkbook;

    private static final String FAILED_CRAWLING_COLUMN = "P";
    private static final int FAILED_CRAWLING_START_ROW = 7;

    private static final String FALSE_NEGATIVE_COLUMN = "Q";
    private static final int FALSE_NEGATIVE_START_ROW = 7;

    private static final String TRUE_NEGATIVE_COLUMN = "R";
    private static final int TRUE_NEGATIVE_START_ROW = 7;

    private static final String EXPERIMENTAL_COLUMN = "S";
    private static final int EXPERIMENTAL_START_ROW = 7;

    private static final String TEST_CASES_COLUMN = "B";
    private static final int TEST_CASES_START_ROW = 7;

    private static final String TYPE_TRUE_POSITIVE_COLUMN = "C";
    private static final String TYPE_FALSE_POSITIVE_COLUMN = "D";
    private static final String TYPE_EXPERIMENTAL_COLUMN = "E";

    private static final short TYPE_TRUE_POSITIVE = 11;
    private static final short TYPE_FALSE_POSITIVE = 22;
    private static final short TYPE_EXPERIMENTAL = 33;

    private static final String GENERATING_TIME_COLUMN = "B";
    private static final int GENERATING_TIME_ROW = 1;

    private static final String DEFAULT_WAVSEP_TEMPLATE_RESOURCE_FILE_PATH = "Template/wavsepTemplate.xlsx";

    public WavsepWriter(String wavsepTemplateFilePath) throws Exception {
        this.wavsepTemplateFilePath = wavsepTemplateFilePath;
        this.wavsepWorkbook = new TcWorkbook(FileUtil.readExcelFile(wavsepTemplateFilePath));
    }

    public WavsepWriter() throws Exception {
        XSSFWorkbook readWorkbook = new XSSFWorkbook(WavsepWriter.class.getClassLoader().getResourceAsStream(DEFAULT_WAVSEP_TEMPLATE_RESOURCE_FILE_PATH));
        this.wavsepTemplateFilePath = DEFAULT_WAVSEP_TEMPLATE_RESOURCE_FILE_PATH;
        this.wavsepWorkbook = new TcWorkbook(readWorkbook);
    }

    /**
     * write failed list in a vulnerability sheet - failed crawling, FALSE Negative, TRUE Negative, Failed Experimental
     *
     * @param vulnerabilityName vulnerability sheet name
     * @param crawledFilePath crawled file path
     * @param scannedFilePath scanned file path
     */
    public void writeFailedListInSheet(String vulnerabilityName, String crawledFilePath, String scannedFilePath){
        TcSheet vulnerabilitySheet = this.wavsepWorkbook.getTcSheet(vulnerabilityName);
        Map<String, Short> testCases = readTestCasesColumn(vulnerabilitySheet);

        setWritingTimeInSheet(vulnerabilitySheet);

        List<String> failedCrawlingList = getFailedCrawlingList(testCases, crawledFilePath);
        List<String> falseNegativeList = getFalseNegativeList(testCases, scannedFilePath);
        List<String> trueNegativeList = getTrueNegativeList(testCases, scannedFilePath);
        List<String> failedScanningExperimentalList = getFailedScanningExperimentalList(testCases, scannedFilePath);

        int startRow = 7;
        String testCasesColumn = "B";
        Cell testCaseCell;
        CellStyle cellStyle = new CellStylesUtil(this.wavsepWorkbook.getWorkbook()).getSimpleCellStyle(false, true, 0, 0, 1, 1);

        for(int i = startRow; (testCaseCell = vulnerabilitySheet.getCell(i, testCasesColumn)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            String testCaseName = testCaseCell.getStringCellValue();
            Cell crawlingCell;
            Cell detectingCell = null;

            // url crawl
            if(failedCrawlingList.contains(testCaseName)){
                crawlingCell = vulnerabilitySheet.getCell(i, "G");
                crawlingCell.setCellValue(false);
            } else {
                crawlingCell = vulnerabilitySheet.getCell(i, "F");
                crawlingCell.setCellValue(true);
            }

            crawlingCell.setCellStyle(cellStyle);

            if(vulnerabilitySheet.getCell(i, "C").getCellType() == Cell.CELL_TYPE_BOOLEAN && vulnerabilitySheet.getCell(i, "C").getBooleanCellValue()){
                if(falseNegativeList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, "I");
                    detectingCell.setCellValue(false);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, "H");
                    detectingCell.setCellValue(true);
                }
            } else if (vulnerabilitySheet.getCell(i, "D").getCellType() == Cell.CELL_TYPE_BOOLEAN && !vulnerabilitySheet.getCell(i, "D").getBooleanCellValue()){
                if(trueNegativeList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, "J");
                    detectingCell.setCellValue(true);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, "K");
                    detectingCell.setCellValue(false);
                }
            } else if (vulnerabilitySheet.getCell(i, "E").getCellType() == Cell.CELL_TYPE_STRING && vulnerabilitySheet.getCell(i, "E").getStringCellValue().equals("EX")){
                if(failedScanningExperimentalList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, "M");
                    detectingCell.setCellValue(false);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, "L");
                    detectingCell.setCellValue(true);
                }
            }
            detectingCell.setCellStyle(cellStyle);
        }
    }

    /**
     * set failed list writing time in a vulnerability sheet
     *
     * @param vulnerabilitySheet vulnerability sheet
     */
    private void setWritingTimeInSheet(TcSheet vulnerabilitySheet){
        vulnerabilitySheet.getCell(GENERATING_TIME_ROW, GENERATING_TIME_COLUMN).setCellValue(new Date());
    }

    /**
     * write all failed list in workbook
     *
     * @param scanResultDirectoryPath directory path that has scan result files
     */
    public void writeAllFailedList(String scanResultDirectoryPath){
        if(scanResultDirectoryPath == null){
            System.err.println("WavsepWriter : Writing all failed list can't do since scanResultDirectory path is null");
            return;
        }

        File scanResultDirectory = new File(scanResultDirectoryPath);

        if(scanResultDirectory.exists() && scanResultDirectory.isDirectory()){
            for(Constant.WavsepSheets wavsepSheet : Constant.WavsepSheets.values()) {
                String crawledFilePath = null;
                String scannedFilePath = null;

                for (File scanResultFile : scanResultDirectory.listFiles()) {
                    if (scanResultFile.getName().contains(wavsepSheet.toString())) {
                        if (crawledFilePath == null && scanResultFile.getName().contains(WavsepParser.WAVSEP_TEST_CRAWLED)) {
                            crawledFilePath = scanResultFile.getAbsolutePath();
                        } else {
                            scannedFilePath = scanResultFile.getAbsolutePath();
                        }
                    }
                }
                if(!wavsepSheet.toString().equals("total")) {
                    writeFailedListInSheet(wavsepSheet.getSheetName(), crawledFilePath, scannedFilePath);
                }
            }
        }
    }

    /**
     * get failed crawling list in a vulnerability sheet
     *
     * @param testCases test cases map
     * @param crawledFilePath crawled file path
     * @return failed crawling list
     */
    private List<String> getFailedCrawlingList(Map<String, Short> testCases, String crawledFilePath) {
        List<String> expectedCrawlingList = new ArrayList<>();
        expectedCrawlingList.addAll(testCases.keySet());
        return WavsepParser.getFailedCrawlingList(crawledFilePath, expectedCrawlingList);
    }

    /**
     * write down false negative list( failed scanning true positive list ) in a vulnerability sheet
     *
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     * @return false negative list
     */
    private List<String> getFalseNegativeList(Map<String, Short> testCases, String scannedFilePath){
        List<String> expectedFalseNegativeList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_TRUE_POSITIVE){
                expectedFalseNegativeList.add(key);
            }
        }
        return WavsepParser.getFalseNegativeList(scannedFilePath, expectedFalseNegativeList);
    }

    /**
     * get true negative list ( failed scanning false positive list ) in a vulnerability sheet
     *
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     * @return true negative list
     */
    private List<String> getTrueNegativeList(Map<String, Short> testCases, String scannedFilePath){
        List<String> expectedTrueNegativeList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_FALSE_POSITIVE){
                expectedTrueNegativeList.add(key);
            }
        }
        return WavsepParser.getTrueNegativeList(scannedFilePath, expectedTrueNegativeList);
//        List<String> trueNegativeList = WavsepParser.getTrueNegativeList(scannedFilePath, expectedTrueNegativeList);
//        TcUtil.writeDownListInSheet(trueNegativeList, vulnerabilitySheet, TRUE_NEGATIVE_START_ROW, TRUE_NEGATIVE_COLUMN);
    }

    /**
     * get failed scanning experimental list
     *
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     * @return failed scanning experimental list
     */
    private List<String> getFailedScanningExperimentalList(Map<String, Short> testCases, String scannedFilePath){
        List<String> expectedScanningExperimentalList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_EXPERIMENTAL){
                expectedScanningExperimentalList.add(key);
            }
        }
        return WavsepParser.getExperimentalList(scannedFilePath, expectedScanningExperimentalList);
    }

    /**
     * read test cases column
     *
     * @param vulnerabilitySheet vulnerability sheet
     * @return test cases map - key : teat case name, value - type short
     */
    private Map<String, Short> readTestCasesColumn(TcSheet vulnerabilitySheet){
        Map<String, Short> testCases = new HashMap<>();

        int testCaseRow = TEST_CASES_START_ROW;
        Cell testCaseCell;

        while((testCaseCell = vulnerabilitySheet.getCell(testCaseRow, TEST_CASES_COLUMN)).getCellType() != Cell.CELL_TYPE_BLANK){
            Cell typeCell;
            if((typeCell = vulnerabilitySheet.getCell(testCaseRow, TYPE_TRUE_POSITIVE_COLUMN)).getCellType() == Cell.CELL_TYPE_BOOLEAN && typeCell.getBooleanCellValue()){
                testCases.put(testCaseCell.getStringCellValue(), TYPE_TRUE_POSITIVE);
            } else if ((typeCell = vulnerabilitySheet.getCell(testCaseRow, TYPE_FALSE_POSITIVE_COLUMN)).getCellType() == Cell.CELL_TYPE_BOOLEAN && !typeCell.getBooleanCellValue()){
                testCases.put(testCaseCell.getStringCellValue(), TYPE_FALSE_POSITIVE);
            } else if((typeCell = vulnerabilitySheet.getCell(testCaseRow, TYPE_EXPERIMENTAL_COLUMN)).getCellType() == Cell.CELL_TYPE_STRING && typeCell.getStringCellValue().equals("EX")){
                testCases.put(testCaseCell.getStringCellValue(), TYPE_EXPERIMENTAL);
            }
            testCaseRow++;
        }

        return testCases;
    }

    /**
     * write excel file
     *
     * @param parentDirectoryPath parent directory path
     * @param fileName file name
     */
    public void writeExcelFile(String parentDirectoryPath, String fileName){
        FormulaEvaluator formulaEvaluator = wavsepWorkbook.getWorkbook().getCreationHelper().createFormulaEvaluator();

        for(Sheet sheet : wavsepWorkbook.getWorkbook()){
            for(Row row : sheet){
                for(Cell cell : row){
                    if(cell.getCellType() == Cell.CELL_TYPE_FORMULA){
                        formulaEvaluator.evaluateFormulaCell(cell);
                    }
                }
            }
        }

        wavsepWorkbook.writeWorkbook(parentDirectoryPath, fileName);
    }

    public static void main(String[] args){
        try {
            Date startDate = new Date();
            WavsepWriter wavsepWriter = new WavsepWriter("C:\\scalaProjects\\testPOI\\wavsepTemplate.xlsx");
            wavsepWriter.writeAllFailedList("C:\\gitProjects\\zap\\results\\160623160427_wavseptest");
            wavsepWriter.writeExcelFile("C:\\Users\\Hwan\\Desktop", "test.xlsx");
            Date endDate = new Date();

            System.out.println("time : " + (endDate.getTime() - startDate.getTime()));

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
