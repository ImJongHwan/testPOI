package org.poi.BenchmarkPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.Constant;
import org.poi.POIConstant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.CellStylesUtil;
import org.poi.Util.FileUtil;
import org.poi.Util.TcUtil;

import java.io.File;
import java.util.*;

/**
 * Created by Hwan on 2016-06-20.
 */
public class BenchmarkWriter {
    private String benchmarkTemplateFilePath;
    private TcWorkbook benchmarkWorkbook;

    private static final int CONTENTS_START_ROW = 7;

    private static final String DEFAULT_BENCHMARK_TEMPLATE_RESOURCE_FILE_PATH = "Template/benchmarkTemplate.xlsx";

    private static final String GENERATING_TIME_COLUMN = "B";
    private static final int GENERATING_TIME_ROW = 1;


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
     * write failed crawling list in Vulnerability Sheet
     *
     * @param vulnerabilityName vulnerability name
     * @param crawledList pure crawled list, not parsing list
     */
    public void writeVulnerabilitySheetContents(String vulnerabilityName, List<String> crawledList, List<String> detectedList){
        TcSheet vulnerabilitySheet = this.benchmarkWorkbook.getTcSheet(vulnerabilityName);
        Map<String, Boolean> testCases = readTestCasesColumn(vulnerabilitySheet);

        List<String> failedCrawlingList = getFailedCrawlingList(testCases, crawledList);
        List<String> falseNegativeList = getFalseNegativeList(testCases, detectedList);
        List<String> trueNegativeList = getTrueNegativeList(testCases, detectedList);

        String testCasesColumn = Constant.TEST_CASES_COLUMN;
        Cell testCaseCell;
        CellStyle cellStyle = new CellStylesUtil(this.benchmarkWorkbook.getWorkbook()).getSimpleCellStyle(false, true, 0, 0, 1, 1);

        setWritingTimeInSheet(vulnerabilitySheet);

        for(int i = CONTENTS_START_ROW; (testCaseCell = vulnerabilitySheet.getCell(i, testCasesColumn)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            String testCaseName = testCaseCell.getStringCellValue();
            Cell crawlingCell;
            Cell detectingCell = null;

            // url crawl
            if(failedCrawlingList.contains(testCaseName)){
                crawlingCell = vulnerabilitySheet.getCell(i, Constant.BenchmarkCrawlColumn.urlCrawlFalse.getColumn());
                crawlingCell.setCellValue(false);
            } else {
                crawlingCell = vulnerabilitySheet.getCell(i, Constant.BenchmarkCrawlColumn.urlCrawlTrue.getColumn());
                crawlingCell.setCellValue(true);
            }

            crawlingCell.setCellStyle(cellStyle);

            // detected true vulnerability
            if(vulnerabilitySheet.getCell(i, Constant.BenchmarkRealVulnerabilityColumn.trueVulnerability.getColumn()).getCellType() == Cell.CELL_TYPE_BOOLEAN
                    && vulnerabilitySheet.getCell(i, Constant.BenchmarkRealVulnerabilityColumn.trueVulnerability.getColumn()).getBooleanCellValue()){
                if(falseNegativeList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.BenchmarkDetectedColumn.detectedTrueVulnerabilityFalse.getColumn());
                    detectingCell.setCellValue(false);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.BenchmarkDetectedColumn.detectedTrueVulnerabilityTrue.getColumn());
                    detectingCell.setCellValue(true);
                }
                // detected false vulnerability
            } else if (vulnerabilitySheet.getCell(i, Constant.BenchmarkRealVulnerabilityColumn.falseVulnerability.getColumn()).getCellType() == Cell.CELL_TYPE_BOOLEAN
                    && !vulnerabilitySheet.getCell(i, Constant.BenchmarkRealVulnerabilityColumn.falseVulnerability.getColumn()).getBooleanCellValue()){
                if(trueNegativeList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.BenchmarkDetectedColumn.detectedFalseVulnerabilityTrue.getColumn());
                    detectingCell.setCellValue(true);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.BenchmarkDetectedColumn.detectedFalseVulnerabilityFalse.getColumn());
                    detectingCell.setCellValue(false);
                }
            }
            detectingCell.setCellStyle(cellStyle);
        }
    }

    /**
     * set time writing data in sheet
     *
     * @param vulnerabilitySheet vulnerability sheet
     */
    private void setWritingTimeInSheet(TcSheet vulnerabilitySheet){
        vulnerabilitySheet.getCell(GENERATING_TIME_ROW, GENERATING_TIME_COLUMN).setCellValue(new Date());
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
                if(!benchmarkSheet.toString().equals(Constant.BenchmarkSheets.total.getSheetName())) {
                    writeVulnerabilitySheetContents(benchmarkSheet.getSheetName(), FileUtil.readFile(crawledFilePath), FileUtil.readFile(scannedFilePath));
                }
            }
        }

    }

    /**
     * get failed crawling list
     *
     * @param testCases test cases map
     * @param pureCrawledList pure crawled list, not parsing list
     * @return failed crawling list
     */
    private List<String> getFailedCrawlingList(Map<String, Boolean> testCases, List<String> pureCrawledList){
        List<String> expectedCrawlingList = new ArrayList<>();
        expectedCrawlingList.addAll(testCases.keySet());

        return BenchmarkParser.getFailedCrawlingList(pureCrawledList, expectedCrawlingList);
    }

    /**
     * get FALSE negative list ( failed scanning true positive list ) in a vulnerability sheet
     *
     * @param testCases test cases map
     * @param pureScannedList pure scanned list, not parsing
     * @return false negative list
     */
    private List<String> getFalseNegativeList(Map<String, Boolean> testCases, List<String> pureScannedList){
        List<String> expectedFalseNegativeList = new ArrayList<>();
        for(String testCase : testCases.keySet()){
            if(testCases.get(testCase)){
                expectedFalseNegativeList.add(testCase);
            }
        }
        return BenchmarkParser.getFalseNegativeList(pureScannedList, expectedFalseNegativeList);
    }

    /**
     * get TRUE negative list ( failed scanning false positive list ) in a vulnerability sheet
     *
     * @param testCases test cases map
     * @param pureScannedList pure scanned list
     * @return true negative list
     */
    private List<String> getTrueNegativeList(Map<String, Boolean> testCases, List<String> pureScannedList){
        List<String> expectedScannedTrueNegativeList = new ArrayList<>();
        for(String testCase : testCases.keySet()){
            if(!testCases.get(testCase)){
                expectedScannedTrueNegativeList.add(testCase);
            }
        }
        return BenchmarkParser.getTrueNegativeList(pureScannedList, expectedScannedTrueNegativeList);
    }

    /**
     * read test cases column
     *
     * @param vulnerabilitySheet vulnerability sheet - TcSheet
     * @return test cases map , key - test case name, value - real vulnerability boolean
     */
    private Map<String, Boolean> readTestCasesColumn(TcSheet vulnerabilitySheet){
        Map<String, Boolean> testCases = new HashMap<>();

        int testCaseRow = CONTENTS_START_ROW;
        Cell testCaseCell;

        while((testCaseCell = vulnerabilitySheet.getCell(testCaseRow, Constant.TEST_CASES_COLUMN)).getCellType() != Cell.CELL_TYPE_BLANK){
            Cell realVulnerabilityTrueCell = vulnerabilitySheet.getCell(testCaseRow, Constant.BenchmarkRealVulnerabilityColumn.trueVulnerability.getColumn());
            if(realVulnerabilityTrueCell.getCellType() == Cell.CELL_TYPE_BOOLEAN && realVulnerabilityTrueCell.getBooleanCellValue()){
                testCases.put(testCaseCell.getStringCellValue(), realVulnerabilityTrueCell.getBooleanCellValue());
            } else {
                testCases.put(testCaseCell.getStringCellValue(), vulnerabilitySheet.getCell(testCaseRow, Constant.BenchmarkRealVulnerabilityColumn.falseVulnerability.getColumn()).getBooleanCellValue());
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
        TcUtil.writeTcWorkbook(parentDirectoryPath, fileName, this.benchmarkWorkbook);
    }

    /**
     * get benchmark template file path
     *
     * @return benchmark template file path
     */
    public String getBenchmarkTemplateFilePath() {
        return benchmarkTemplateFilePath;
    }

    /**
     * get benchmark TcWorkbook
     *
     * @return benchmark TcWorkbook
     */
    public TcWorkbook getBenchmarkWorkbook() {
        return benchmarkWorkbook;
    }

    /**
     * testing
     * @param args empty
     */
    public static void main(String[] args){
        try {
            Date startDate = new Date();
            BenchmarkWriter benchmarkWriter = new BenchmarkWriter("C:\\scalaProjects\\testPOI\\src\\main\\resources\\Template\\benchmarkTemplate.xlsx");
            benchmarkWriter.writeAllFailedList("C:\\gitProjects\\zap\\results\\160627144327_benchmarktest");
            benchmarkWriter.writeExcelFile("C:\\Users\\Hwan\\Desktop", "test" + POIConstant.EXCEL_FILE_EXTENSION);
            Date endDate = new Date();

            System.out.println("time : " + (endDate.getTime() - startDate.getTime()));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}