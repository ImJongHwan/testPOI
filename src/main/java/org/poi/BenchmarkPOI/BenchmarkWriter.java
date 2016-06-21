package org.poi.BenchmarkPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.Constant;
import org.poi.POIConstant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.FileUtil;
import org.poi.Util.TcUtil;

import java.io.File;
import java.io.IOException;
import java.util.*;

/**
 * Created by Hwan on 2016-06-20.
 */
public class BenchmarkWriter {
    private String benchmarkTemplateFilePath;
    private TcWorkbook benchmarkWorkbook;

    private static final String FAILED_CRAWLING_START_COLUMN = "M";
    private static final int FAILED_CRAWLING_START_ROW = 7;

    private static final String FALSE_NEGATIVE_START_COLUMN = "N";
    private static final int FALSE_NEGATIVE_START_ROW = 7;

    private static final String TRUE_NEGATIVE_START_COLUMN = "O";
    private static final int TRUE_NEGATIVE_START_ROW = 7;

    private static final String TEST_CASES_COLUMN = "B";
    private static final int TEST_CASES_START_ROW = 7;

    private static final String REAL_VULNERABILITY_TRUE_COLUMN = "C";
    private static final String REAL_VULNERABILITY_FALSE_COLUMN = "D";

    private static final String DEFAULT_BENCHMARK_TEMPLATE_RESOURCE_FILE_PATH = "Template/benchmarkTemplate.xlsx";


    public BenchmarkWriter(String benchmarkTemplateFilePath) throws Exception {
        this.benchmarkTemplateFilePath = benchmarkTemplateFilePath;
        this.benchmarkWorkbook = new TcWorkbook(FileUtil.readExcelFile(benchmarkTemplateFilePath));
    }

    public BenchmarkWriter() throws Exception {
        XSSFWorkbook readWorkbook = new XSSFWorkbook(BenchmarkWriter.class.getClassLoader().getResourceAsStream(DEFAULT_BENCHMARK_TEMPLATE_RESOURCE_FILE_PATH));
        this.benchmarkTemplateFilePath = DEFAULT_BENCHMARK_TEMPLATE_RESOURCE_FILE_PATH;
        this.benchmarkWorkbook = new TcWorkbook(readWorkbook);
    }

    /**
     * write failed list in a vulnerability sheet - Failed Crwaling, FALSE Negative, TRUE Negative
     *
     * @param vulnerabilityName vulnerability Name
     * @param crawledFilePath crawled file path
     * @param scannedFilePath scanned file path
     * @throws IOException
     */
    public void writeFailedListInSheet(String vulnerabilityName, String crawledFilePath, String scannedFilePath) throws IOException {
        TcSheet vulnerabilitySheet = this.benchmarkWorkbook.getTcSheet(vulnerabilityName);
        Map<String, Boolean> testCases = readTestCasesColumn(vulnerabilitySheet);

        writeFailedCrawlingList(vulnerabilitySheet, testCases, crawledFilePath);
        writeFailedScanningFalseNegativeList(vulnerabilitySheet, testCases, scannedFilePath);
        writeFailedScanningTrueNegativeList(vulnerabilitySheet, testCases, scannedFilePath);
    }

    /**
     * write all failed list in workbook
     *
     * @param scanResultDirectoryPath scan result directory path
     */
    public void writeAllFailedList(String scanResultDirectoryPath){
        if(scanResultDirectoryPath == null){
            System.err.println("BenchmarkWriter : Writing all failed list can't do since scanResultDirectory path is null");
            return;
        }

        File scanResultDirectory = new File(scanResultDirectoryPath);

        if(scanResultDirectory.exists() && scanResultDirectory.isDirectory()){
            for(Constant.BenchmarkSheets benchmarkSheet : Constant.BenchmarkSheets.values()){
                try {
                    String crawledFilePath = null;
                    String scannedFilePath = null;
                    for (File scanResultFile : scanResultDirectory.listFiles()) {
                        if (scanResultFile.getName().contains(benchmarkSheet.toString())) {
                            if (crawledFilePath == null && scanResultFile.getName().contains(BenchmarkParser.BENCHMARK_TEST_CRAWLED)) {
                                crawledFilePath = scanResultFile.getAbsolutePath();
                            } else {
                                scannedFilePath = scanResultFile.getAbsolutePath();
                            }
                        }
                    }
                    if (crawledFilePath != null && scannedFilePath != null) {
                        writeFailedListInSheet(benchmarkSheet.getSheetName(), crawledFilePath, scannedFilePath);
                    }
                } catch (IOException e){
                    System.err.println("BenchmarkWriter : Writing failed list in sheet is failed since occur IOException - " + benchmarkSheet.getSheetName());
                    e.printStackTrace();
                }
            }
        }

    }


    /**
     * write down failed crawling list
     *
     * @param vulnerabilitySheet vulnerability sheet
     * @param testCases test cases map
     * @param crawledFilePath crawled file path
     * @throws IOException
     */
    private void writeFailedCrawlingList(TcSheet vulnerabilitySheet, Map<String, Boolean> testCases, String crawledFilePath) throws IOException {
        List<String> expectedCrawlingList = new ArrayList<>();
        expectedCrawlingList.addAll(testCases.keySet());
        List<String> failedCrawlingList = BenchmarkParser.getFailedCrawlingList(crawledFilePath, expectedCrawlingList);
        TcUtil.writeDownListInSheet(failedCrawlingList, vulnerabilitySheet, FAILED_CRAWLING_START_ROW, FAILED_CRAWLING_START_COLUMN);
    }

    /**
     * write down failed scanning false negative list
     *
     * @param vulnerabilitySheet vulnerability sheet
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     * @throws IOException
     */
    private void writeFailedScanningFalseNegativeList(TcSheet vulnerabilitySheet, Map<String, Boolean> testCases, String scannedFilePath) throws IOException {
        List<String> expectedScannedFalseNegativeList = new ArrayList<>();
        for(String testCase : testCases.keySet()){
            if(testCases.get(testCase)){
                expectedScannedFalseNegativeList.add(testCase);
            }
        }
        List<String> failedScanningFalseNegativeList = BenchmarkParser.getFalseNegativeList(scannedFilePath, expectedScannedFalseNegativeList);
        TcUtil.writeDownListInSheet(failedScanningFalseNegativeList, vulnerabilitySheet, FALSE_NEGATIVE_START_ROW, FALSE_NEGATIVE_START_COLUMN);
    }

    /**
     * write down failed scanning true negative list
     *
     * @param vulnerabilitySheet vulnerability sheet
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     * @throws IOException
     */
    private void writeFailedScanningTrueNegativeList(TcSheet vulnerabilitySheet, Map<String, Boolean> testCases, String scannedFilePath) throws IOException {
        List<String> expectedScannedTrueNegativeList = new ArrayList<>();
        for(String testCase : testCases.keySet()){
            if(!testCases.get(testCase)){
                expectedScannedTrueNegativeList.add(testCase);
            }
        }
        List<String> failedScanningTrueNegativeList = BenchmarkParser.getFalsePositiveList(scannedFilePath, expectedScannedTrueNegativeList);
        TcUtil.writeDownListInSheet(failedScanningTrueNegativeList, vulnerabilitySheet, TRUE_NEGATIVE_START_ROW, TRUE_NEGATIVE_START_COLUMN);
    }

    /**
     * read test cases column
     *
     * @param vulnerabilitySheet vulnerability sheet - TcSheet
     * @return test cases map , key - test case name, value - real vulnerability boolean
     */
    private Map<String, Boolean> readTestCasesColumn(TcSheet vulnerabilitySheet){
        Map<String, Boolean> testCases = new HashMap<>();

        int testCaseRow = TEST_CASES_START_ROW;
        Cell testCaseCell;

        while((testCaseCell = vulnerabilitySheet.getCell(testCaseRow, TEST_CASES_COLUMN)).getCellType() != Cell.CELL_TYPE_BLANK){
            Cell realVulnerabilityTrueCell = vulnerabilitySheet.getCell(testCaseRow, REAL_VULNERABILITY_TRUE_COLUMN);
            if(realVulnerabilityTrueCell.getCellType() == Cell.CELL_TYPE_BOOLEAN && realVulnerabilityTrueCell.getBooleanCellValue()){
                testCases.put(testCaseCell.getStringCellValue(), realVulnerabilityTrueCell.getBooleanCellValue());
            } else {
                testCases.put(testCaseCell.getStringCellValue(), vulnerabilitySheet.getCell(testCaseRow, REAL_VULNERABILITY_FALSE_COLUMN).getBooleanCellValue());
            }
            testCaseRow++;
        }

        return testCases;
    }

    /**
     * write this workbook in file
     *
     * @param parentDirectoryPath parent directory path
     * @param fileName file name
     */
    public void writeExcelFile(String parentDirectoryPath, String fileName){
        benchmarkWorkbook.getWorkbook().setForceFormulaRecalculation(true);
        benchmarkWorkbook.writeWorkbook(parentDirectoryPath, fileName);
    }

    public String getBenchmarkTemplateFilePath() {
        return benchmarkTemplateFilePath;
    }

    public TcWorkbook getBenchmarkWorkbook() {
        return benchmarkWorkbook;
    }

    public static void main(String[] args){
        try {
            BenchmarkWriter benchmarkWriter = new BenchmarkWriter();
            benchmarkWriter.writeAllFailedList("C:\\gitProjects\\zap\\results\\160620141843_benchmarktest");
            benchmarkWriter.writeExcelFile("C:\\Users\\Hwan\\Desktop", "test" + POIConstant.EXCEL_FILE_EXTENSION);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}