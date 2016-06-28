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

    private static final String FAILED_CRAWLING_START_COLUMN = "M";
    private static final int FAILED_CRAWLING_START_ROW = 7;

    private static final String FALSE_NEGATIVE_COLUMN = "N";
    private static final int FALSE_NEGATIVE_START_ROW = 7;

    private static final String TRUE_NEGATIVE_COLUMN = "O";
    private static final int TRUE_NEGATIVE_START_ROW = 7;

    private static final String TEST_CASES_COLUMN = "B";
    private static final int TEST_CASES_START_ROW = 7;

    private static final String REAL_VULNERABILITY_TRUE_COLUMN = "C";
    private static final String REAL_VULNERABILITY_FALSE_COLUMN = "D";

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
     * write failed list in a vulnerability sheet - Failed Crwaling, FALSE Negative, TRUE Negative
     *
     * @param vulnerabilityName vulnerability Name
     * @param crawledFilePath crawled file path
     * @param scannedFilePath scanned file path
     */
    public void writeFailedListInSheet(String vulnerabilityName, String crawledFilePath, String scannedFilePath){
        TcSheet vulnerabilitySheet = this.benchmarkWorkbook.getTcSheet(vulnerabilityName);
        Map<String, Boolean> testCases = readTestCasesColumn(vulnerabilitySheet);

        setWritingTimeInSheet(vulnerabilitySheet);

        // get list
        List<String> failedCrawlingList = getFailedCrawlingList(testCases, crawledFilePath);
        List<String> falseNegativeList = getFalseNegativeList(testCases, scannedFilePath);
        List<String> trueNegativeList = getTrueNegativeList(testCases, scannedFilePath);

        int startRow = 7;
        String testCasesColumn = "B";
        Cell testCaseCell;
        CellStyle cellStyle = new CellStylesUtil(this.benchmarkWorkbook.getWorkbook()).getSimpleCellStyle(false, true, 0, 0, 1, 1);

        for(int i = startRow; (testCaseCell = vulnerabilitySheet.getCell(i, testCasesColumn)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            String testCaseName = testCaseCell.getStringCellValue();
            Cell crawlingCell;
            Cell detectingCell = null;

            // url crawl
            if(failedCrawlingList.contains(testCaseName)){
                crawlingCell = vulnerabilitySheet.getCell(i, "F");
                crawlingCell.setCellValue(false);
            } else {
                crawlingCell = vulnerabilitySheet.getCell(i, "E");
                crawlingCell.setCellValue(true);
            }

            crawlingCell.setCellStyle(cellStyle);

            // detected true vulnerability
            if(vulnerabilitySheet.getCell(i, "C").getCellType() == Cell.CELL_TYPE_BOOLEAN && vulnerabilitySheet.getCell(i, "C").getBooleanCellValue()){
                if(falseNegativeList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, "H");
                    detectingCell.setCellValue(false);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, "G");
                    detectingCell.setCellValue(true);
                }
                // detected false vulnerability
            } else if (vulnerabilitySheet.getCell(i, "D").getCellType() == Cell.CELL_TYPE_BOOLEAN && !vulnerabilitySheet.getCell(i, "D").getBooleanCellValue()){
                if(trueNegativeList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, "I");
                    detectingCell.setCellValue(true);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, "J");
                    detectingCell.setCellValue(false);
                }
            }
            detectingCell.setCellStyle(cellStyle);
        }
    }

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
                if(!benchmarkSheet.toString().equals("total")) {
                    writeFailedListInSheet(benchmarkSheet.getSheetName(), crawledFilePath, scannedFilePath);
                }
            }
        }

    }


    /**
     * get failed crawling list
     *
     * @param testCases test cases map
     * @param crawledFilePath crawled file path
     * @return failed crawling list
     */
    private List<String> getFailedCrawlingList(Map<String, Boolean> testCases, String crawledFilePath){
        List<String> expectedCrawlingList = new ArrayList<>();
        expectedCrawlingList.addAll(testCases.keySet());

        return BenchmarkParser.getFailedCrawlingList(crawledFilePath, expectedCrawlingList);
    }

    /**
     * get false negative list(failed scanning true positive list)
     *
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     * @return false negative list
     */
    private List<String> getFalseNegativeList(Map<String, Boolean> testCases, String scannedFilePath){
        List<String> expectedFalseNegativeList = new ArrayList<>();
        for(String testCase : testCases.keySet()){
            if(testCases.get(testCase)){
                expectedFalseNegativeList.add(testCase);
            }
        }
        return BenchmarkParser.getFalseNegativeList(scannedFilePath, expectedFalseNegativeList);
    }

    /**
     * get true negative list( failed scanning false positive list )
     *
     * @param testCases test cases map
     * @param scannedFilePath scanned file path
     * @return true negative list
     */
    private List<String> getTrueNegativeList(Map<String, Boolean> testCases, String scannedFilePath){
        List<String> expectedScannedTrueNegativeList = new ArrayList<>();
        for(String testCase : testCases.keySet()){
            if(!testCases.get(testCase)){
                expectedScannedTrueNegativeList.add(testCase);
            }
        }
        return BenchmarkParser.getFalsePositiveList(scannedFilePath, expectedScannedTrueNegativeList);
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
        TcUtil.writeTcWorkbook(parentDirectoryPath, fileName, this.benchmarkWorkbook);
    }

    public String getBenchmarkTemplateFilePath() {
        return benchmarkTemplateFilePath;
    }

    public TcWorkbook getBenchmarkWorkbook() {
        return benchmarkWorkbook;
    }

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