package org.poi.BenchmarkPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.Constant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.FileUtil;
import org.poi.Util.TcUtil;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Created by Hwan on 2016-06-07.
 */
public class BenchmarkComparer {

    private String benchmarkCompareTemplateFilePath;
    private TcWorkbook benchmarkCompareWorkbook;

    private static final int VULNERABILITY_CONTENTS_START_ROW = 7;
    private static final String TEST_CASE_COLUMN = "B";

    private static final int COMPARE_VULNERABILITY_START_ROW = 6;

    private static final String DEFAULT_BENCHMARK_COMPARE_TEMPLATE_RESOURCE_FILE_PATH = "Template/benchmarkCompareTemplate.xlsx";

    public BenchmarkComparer(String benchmarkCompareTemplateFilePath) throws IOException {
        this.benchmarkCompareTemplateFilePath = benchmarkCompareTemplateFilePath;
        this.benchmarkCompareWorkbook = new TcWorkbook(FileUtil.readExcelFile(benchmarkCompareTemplateFilePath));
    }

    public BenchmarkComparer() throws IOException{
        this.benchmarkCompareTemplateFilePath = DEFAULT_BENCHMARK_COMPARE_TEMPLATE_RESOURCE_FILE_PATH;
        XSSFWorkbook readWorkbook = new XSSFWorkbook(BenchmarkComparer.class.getClassLoader().getResourceAsStream(DEFAULT_BENCHMARK_COMPARE_TEMPLATE_RESOURCE_FILE_PATH));
        this.benchmarkCompareWorkbook = new TcWorkbook(readWorkbook);
    }

    /**
     * write compare contents in a benchmark compare workbook
     *
     * @param beforeWorkbookPath before workbook file path
     * @param afterWorkbookPath after workbook file path
     * @throws IOException
     */
    public void writeCompareInWorkbook(String beforeWorkbookPath, String afterWorkbookPath) throws IOException {
        TcWorkbook beforeBenchmarkWorkbook = new TcWorkbook(FileUtil.readExcelFile(beforeWorkbookPath));
        TcWorkbook afterBenchmarkWorkbook = new TcWorkbook(FileUtil.readExcelFile(afterWorkbookPath));

        for(Constant.BenchmarkSheets benchmarkSheet : Constant.BenchmarkSheets.values()){
            String benchmarkSheetName = benchmarkSheet.getSheetName();
            if(benchmarkSheet.toString().equals("total")){
                setBeforeAndAfterWorkbookTitleInTotalSheet(benchmarkCompareWorkbook.getTcSheet(benchmarkSheetName), beforeWorkbookPath, afterWorkbookPath);
                writeExistingTotalsInTotalSheet(benchmarkCompareWorkbook.getTcSheet(benchmarkSheetName), beforeBenchmarkWorkbook.getTcSheet(benchmarkSheetName), afterBenchmarkWorkbook.getTcSheet(benchmarkSheetName));
            } else {
                writeCompareInVulnerabilitySheet(benchmarkCompareWorkbook.getTcSheet(benchmarkSheetName), beforeBenchmarkWorkbook.getTcSheet(benchmarkSheetName), afterBenchmarkWorkbook.getTcSheet(benchmarkSheetName));
            }
        }
    }

    /**
     * set before and after workbook title name in a compare total sheet
     *
     * @param compareTotalSheet compare total sheet
     * @param beforeWorkbookPath before workbook file path
     * @param afterWorkbookPath after workbook file path
     */
    private void setBeforeAndAfterWorkbookTitleInTotalSheet(TcSheet compareTotalSheet, String beforeWorkbookPath, String afterWorkbookPath){
        final String BEFORE_TITLE_COLUMN = "A";
        final int BEFORE_TITLE_ROW = 19;

        final String AFTER_TITLE_COLUMN = "A";
        final int AFTER_TITLE_ROW = 36;

        compareTotalSheet.getCell(BEFORE_TITLE_ROW, BEFORE_TITLE_COLUMN).setCellValue(getFileNameFromPath(beforeWorkbookPath) + " (A Case)");
        compareTotalSheet.getCell(AFTER_TITLE_ROW, AFTER_TITLE_COLUMN).setCellValue(getFileNameFromPath(afterWorkbookPath) + " (B Case)");
    }

    /**
     * get file name from file path
     *
     * @param path file path
     * @return file name except for a file extension
     */
    private String getFileNameFromPath(String path){
        String fileName = new File(path).getName();
        return fileName.substring(0, fileName.lastIndexOf("."));
    }

    /**
     * write before and after total sheet existing total contents in compare total sheet
     * @param compareTotalSheet compare total sheet
     * @param beforeTotalSheet before total sheet
     * @param afterTotalSheet after total sheet
     */
    private void writeExistingTotalsInTotalSheet(TcSheet compareTotalSheet, TcSheet beforeTotalSheet, TcSheet afterTotalSheet){
        final int startRowForCopyBefore = 20;
        final int startRowForCopyAfter = 37;

        copyTotalContents(beforeTotalSheet, compareTotalSheet, startRowForCopyBefore);
        copyTotalContents(afterTotalSheet, compareTotalSheet, startRowForCopyAfter);
    }

    /**
     * copy total contents in a compare total sheet
     *
     * @param copiedBenchmarkTotalSheet copied benchmark total sheet
     * @param compareBenchmarkTotalSheet compare benchmark total sheet
     * @param compareTotalRowForCopy compare total row for copy
     */
    private void copyTotalContents(TcSheet copiedBenchmarkTotalSheet, TcSheet compareBenchmarkTotalSheet, int compareTotalRowForCopy){
        final int copiedStartRow = 1;
        final int copiedEndRow = 14;

        //copy rows
        for(int i = copiedStartRow; i<= copiedEndRow; i++){
            Iterator<Cell> cellIterator = copiedBenchmarkTotalSheet.getSheet().getRow(i - 1).iterator();
            while(cellIterator.hasNext()){
                Cell copiedCell = cellIterator.next();
                Cell compareCell = compareBenchmarkTotalSheet.getCell(compareTotalRowForCopy - 1, copiedCell.getColumnIndex());

                // copy cell value
                switch(TcUtil.getCellTypeUnwrappedFormula(copiedCell)){
                    case Cell.CELL_TYPE_BOOLEAN :
                        compareCell.setCellValue(copiedCell.getBooleanCellValue());
                        break;
                    case Cell.CELL_TYPE_NUMERIC :
                        compareCell.setCellValue(copiedCell.getNumericCellValue());
                        break;
                    case Cell.CELL_TYPE_STRING :
                        compareCell.setCellValue(copiedCell.getStringCellValue());
                        break;
                    case Cell.CELL_TYPE_ERROR :
                        compareCell.setCellValue(copiedCell.getErrorCellValue());
                        break;
                }
            }
            compareTotalRowForCopy++;
        }
    }

    /**
     * write compare in compare vulnerability sheet
     *
     * @param compareVulnerabilitySheet compare vulnerability sheet
     * @param beforeVulnerabilitySheet before vulnerability sheet
     * @param afterVulnerabilitySheet after vulnerability sheet
     */
    private void writeCompareInVulnerabilitySheet(TcSheet compareVulnerabilitySheet, TcSheet beforeVulnerabilitySheet, TcSheet afterVulnerabilitySheet){
        Map<String, Boolean> beforeCrawlingMap = readCrawlingTestCases(beforeVulnerabilitySheet);
        Map<String, Boolean> afterCrawlingMap = readCrawlingTestCases(afterVulnerabilitySheet);
        Map<String, Boolean> beforeDetectingMap = readDetectingTestCases(beforeVulnerabilitySheet);
        Map<String, Boolean> afterDetectingMap = readDetectingTestCases(afterVulnerabilitySheet);

        Cell testCaseCell;

        for(int i = COMPARE_VULNERABILITY_START_ROW; (testCaseCell = compareVulnerabilitySheet.getCell(i, Constant.TEST_CASES_COLUMN)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            String testCaseName = testCaseCell.getStringCellValue();
            short flagDiff = 0;

            if(beforeCrawlingMap.get(testCaseName)){
                compareVulnerabilitySheet.getCell(i, "E").setCellValue(true);
            }

            if(afterCrawlingMap.get(testCaseName)){
                compareVulnerabilitySheet.getCell(i, "F").setCellValue(true);
            }

            if(beforeDetectingMap.get(testCaseName)){
                compareVulnerabilitySheet.getCell(i, "G").setCellValue(true);
                flagDiff++;
            }

            if(afterDetectingMap.get(testCaseName)){
                compareVulnerabilitySheet.getCell(i, "H").setCellValue(true);
                flagDiff++;
            }

            if(flagDiff == 1){
                if(beforeDetectingMap.get(testCaseName)){
                    compareVulnerabilitySheet.getCell(i, "I").setCellValue(true);
                } else {
                    compareVulnerabilitySheet.getCell(i, "J").setCellValue(true);
                }
            }
        }
    }

    /**
     * read test cases included crawling result in benchmark vulnerability sheet
     *
     * @param benchmarkVulnerabilitySheet benchmark vulnerability sheet
     * @return crawling test cases map, key - test case name, value - isCrawled boolean
     */
    private Map<String, Boolean> readCrawlingTestCases(TcSheet benchmarkVulnerabilitySheet){
        Map<String, Boolean> crawlingTestCases = new HashMap<>();
        Cell testCaseCell;

        for(int i = VULNERABILITY_CONTENTS_START_ROW; (testCaseCell = benchmarkVulnerabilitySheet.getCell(i, TEST_CASE_COLUMN)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            Cell trueCrawlCell = benchmarkVulnerabilitySheet.getCell(i, Constant.BenchmarkCrawlColumn.urlCrawlTrue.getColumn());
            Cell falseCrawlCell = benchmarkVulnerabilitySheet.getCell(i, Constant.BenchmarkCrawlColumn.urlCrawlFalse.getColumn());

            if(TcUtil.getCellTypeUnwrappedFormula(trueCrawlCell) == Cell.CELL_TYPE_STRING && trueCrawlCell.getStringCellValue().equals("TRUE")){
                crawlingTestCases.put(testCaseCell.getStringCellValue(), Boolean.valueOf(trueCrawlCell.getStringCellValue()));
            } else if (TcUtil.getCellTypeUnwrappedFormula(falseCrawlCell) == Cell.CELL_TYPE_STRING && falseCrawlCell.getStringCellValue().equals("FALSE")){
                crawlingTestCases.put(testCaseCell.getStringCellValue(), Boolean.valueOf(falseCrawlCell.getStringCellValue()));
            }
        }
        return crawlingTestCases;
    }

    /**
     * read test cases included detecting result in benchmark vulnerability sheet
     *
     * @param benchmarkVulnerabilitySheet benchmark vulnerability sheet
     * @return detecting test cases map, key - test case name, value - isDetected boolean
     */
    private Map<String, Boolean> readDetectingTestCases(TcSheet benchmarkVulnerabilitySheet) {
        Map<String, Boolean> detectingTestCases = new HashMap<>();
        Cell testCaseCell;

        for(int i = VULNERABILITY_CONTENTS_START_ROW; (testCaseCell = benchmarkVulnerabilitySheet.getCell(i, TEST_CASE_COLUMN)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            for(Constant.BenchmarkDetectedColumn benchmarkDetectedColumn : Constant.BenchmarkDetectedColumn.values()){
                Cell detectedCell = benchmarkVulnerabilitySheet.getCell(i, benchmarkDetectedColumn.getColumn());
                if(TcUtil.getCellTypeUnwrappedFormula(detectedCell) == Cell.CELL_TYPE_STRING){
                    detectingTestCases.put(testCaseCell.getStringCellValue(), Boolean.valueOf(detectedCell.getStringCellValue()));
                    break;
                }
            }
        }
        return detectingTestCases;
    }

    /**
     * get benchmark compare template file path
     * @return benchmark compare template file path
     */
    public String getBenchmarkCompareTemplateFilePath() {
        return benchmarkCompareTemplateFilePath;
    }

    /**
     * get benchmark compare workbook
     * @return benchmark compare workbook
     */
    public TcWorkbook getBenchmarkCompareWorkbook() {
        return benchmarkCompareWorkbook;
    }

    /**
     * write benchmark compare workbook current state in excel file
     *
     * @param parentDirectoryPath parent directory path
     * @param fileName file name
     */
    public void writeBenchmarkCompareWorkbookExcelFile(String parentDirectoryPath, String fileName){
        TcUtil.writeTcWorkbook(parentDirectoryPath, fileName, this.benchmarkCompareWorkbook);
    }

    public static void main(String[] args){
        try {
            BenchmarkComparer benchmarkComparer = new BenchmarkComparer("C:\\scalaProjects\\testPOI\\src\\main\\resources\\Template\\benchmarkCompareTemplate.xlsx");
            benchmarkComparer.writeCompareInWorkbook("C:\\scalaProjects\\testPOI\\Zap_Benchmark_Results160609153221.xlsx", "Zap_Benchmark_Results160609165623.xlsx");
            benchmarkComparer.writeBenchmarkCompareWorkbookExcelFile("C:\\Users\\Hwan\\Desktop", "test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}