package org.poi.WavsepPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.Constant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
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

        writeFailedCrawlingList(vulnerabilitySheet, testCases, crawledFilePath);
        writeFalseNegativeList(vulnerabilitySheet, testCases, scannedFilePath);
        writeTrueNegativeList(vulnerabilitySheet, testCases, scannedFilePath);
        writeFailedScanningExperimentalList(vulnerabilitySheet, testCases, scannedFilePath);
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
     * write down failed crawling list in a vulnerability sheet
     *
     * @param vulnerabilitySheet vulnerability sheet
     * @param testCases test cases map
     * @param crawledFilePath crawled file path
     */
    private void writeFailedCrawlingList(TcSheet vulnerabilitySheet, Map<String, Short> testCases, String crawledFilePath) {
        List<String> expectedCrawlingList = new ArrayList<>();
        expectedCrawlingList.addAll(testCases.keySet());
        List<String> failedCrawlingList = WavsepParser.getFailedCrawlingList(crawledFilePath, expectedCrawlingList);
        TcUtil.writeDownListInSheet(failedCrawlingList, vulnerabilitySheet, FAILED_CRAWLING_START_ROW, FAILED_CRAWLING_COLUMN);
    }

    /**
     * write down false negative list( failed scanning true positive list ) in a vulnerability sheet
     *
     * @param vulnerabilitySheet vulenrability sheet
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     */
    private void writeFalseNegativeList(TcSheet vulnerabilitySheet, Map<String, Short> testCases, String scannedFilePath){
        List<String> expectedFalseNegativeList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_TRUE_POSITIVE){
                expectedFalseNegativeList.add(key);
            }
        }
        List<String> falseNegativeList = WavsepParser.getFalseNegativeList(scannedFilePath, expectedFalseNegativeList);
        TcUtil.writeDownListInSheet(falseNegativeList, vulnerabilitySheet, FALSE_NEGATIVE_START_ROW, FALSE_NEGATIVE_COLUMN);
    }

    /**
     * write down true negative list ( failed scanning false positive list ) in a vulnerability sheet
     *
     * @param vulnerabilitySheet vulnerability sheet
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     */
    private void writeTrueNegativeList(TcSheet vulnerabilitySheet, Map<String, Short> testCases, String scannedFilePath){
        List<String> expectedTrueNegativeList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_FALSE_POSITIVE){
                expectedTrueNegativeList.add(key);
            }
        }
        List<String> trueNegativeList = WavsepParser.getTrueNegativeList(scannedFilePath, expectedTrueNegativeList);
        TcUtil.writeDownListInSheet(trueNegativeList, vulnerabilitySheet, TRUE_NEGATIVE_START_ROW, TRUE_NEGATIVE_COLUMN);
    }

    /**
     * write down failed scanning experimental list
     *
     * @param vulnerabilitySheet vulnerability sheet
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     */
    private void writeFailedScanningExperimentalList(TcSheet vulnerabilitySheet, Map<String, Short> testCases, String scannedFilePath){
        List<String> expectedScanningExperimentalList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_EXPERIMENTAL){
                expectedScanningExperimentalList.add(key);
            }
        }
        List<String> scannedExperimentalList = WavsepParser.getExperimentalList(scannedFilePath, expectedScanningExperimentalList);
        TcUtil.writeDownListInSheet(scannedExperimentalList, vulnerabilitySheet, EXPERIMENTAL_START_ROW, EXPERIMENTAL_COLUMN);
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
        wavsepWorkbook.writeWorkbook(parentDirectoryPath, fileName);
    }

    public static void main(String[] args){
        try {
            WavsepWriter wavsepWriter = new WavsepWriter("C:\\scalaProjects\\testPOI\\src\\main\\resources\\Template\\wavsepTemplate.xlsx");
            wavsepWriter.writeAllFailedList("C:\\gitProjects\\zap\\results\\160615155330_wavseptest");
            wavsepWriter.writeExcelFile("C:\\Users\\Hwan\\Desktop", "test.xlsx");
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
