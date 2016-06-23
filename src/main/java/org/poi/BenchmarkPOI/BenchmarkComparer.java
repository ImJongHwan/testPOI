package org.poi.BenchmarkPOI;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.Constant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.FileUtil;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Hwan on 2016-06-07.
 */
public class BenchmarkComparer {

    private String benchmarkCompareTemplateFilePath;
    private TcWorkbook benchmarkCompareWorkbook;

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
        final int BEFORE_TITLE_ROW = 17;

        final String AFTER_TITLE_COLUMN = "A";
        final int AFTER_TITLE_ROW = 31;

        compareTotalSheet.getCell(BEFORE_TITLE_ROW, BEFORE_TITLE_COLUMN).setCellValue(getFileNameFromPath(beforeWorkbookPath));
        compareTotalSheet.getCell(AFTER_TITLE_ROW, AFTER_TITLE_COLUMN).setCellValue(getFileNameFromPath(afterWorkbookPath));
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

    }

    /**
     * write compare in compare vulnerability sheet
     *
     * @param compareVulnerabilitySheet compare vulnerability sheet
     * @param beforeVulnerabilitySheet before vulnerability sheet
     * @param afterVulnerabilitySheet after vulnerability sheet
     */
    private void writeCompareInVulnerabilitySheet(TcSheet compareVulnerabilitySheet, TcSheet beforeVulnerabilitySheet, TcSheet afterVulnerabilitySheet){

    }

    /**
     * read test cases included crawling result in benchmark vulnerability sheet
     *
     * @param benchmarkVulnerabilitySheet benchmark vulnerability sheet
     * @return crawling test cases map, key - test case name, value - isCrawled boolean
     */
    private Map<String, Boolean> readCrawlingTestCases(TcSheet benchmarkVulnerabilitySheet){
        Map<String, Boolean> crawlingTestCases = new HashMap<>();
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
}