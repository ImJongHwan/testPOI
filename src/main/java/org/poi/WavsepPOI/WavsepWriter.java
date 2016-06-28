package org.poi.WavsepPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
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

    private static final int TEST_CASES_START_ROW = 7;

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
     * write vulnerability sheet contents
     *
     * @param vulnerabilitySheetName vulnerability sheet name
     * @param crawledList pure crawled list
     * @param detectedList pure detected list
     */
    public void writeVulnerabilitySheetContents(String vulnerabilitySheetName, List<String> crawledList, List<String> detectedList){
        TcSheet vulnerabilitySheet = this.wavsepWorkbook.getTcSheet(vulnerabilitySheetName);
        Map<String, Short> testCases = readTestCasesColumn(vulnerabilitySheet);

        List<String> failedCrawlingList = getFailedCrawlingList(testCases, crawledList);
        List<String> falseNegativeList = getFalseNegativeList(testCases, detectedList);
        List<String> trueNegativeList = getTrueNegativeList(testCases, detectedList);
        List<String> failedScanningExperimentalList = getFailedScanningExperimentalList(testCases, detectedList);

        int startRow = 7;
        Cell testCaseCell;
        CellStyle cellStyle = new CellStylesUtil(this.wavsepWorkbook.getWorkbook()).getSimpleCellStyle(false, true, 0, 0, 1, 1);

        setWritingTimeInSheet(vulnerabilitySheet);

        for(int i = startRow; (testCaseCell = vulnerabilitySheet.getCell(i, Constant.TEST_CASES_COLUMN)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            String testCaseName = testCaseCell.getStringCellValue();
            Cell crawlingCell;
            Cell detectingCell = null;

            // url crawl
            if(failedCrawlingList.contains(testCaseName)){
                crawlingCell = vulnerabilitySheet.getCell(i, Constant.WavsepCrawlColumn.urlCrawlFalse.getColumn());
                crawlingCell.setCellValue(false);
            } else {
                crawlingCell = vulnerabilitySheet.getCell(i, Constant.WavsepCrawlColumn.urlCrawlTrue.getColumn());
                crawlingCell.setCellValue(true);
            }

            crawlingCell.setCellStyle(cellStyle);

            // detected true positive type
            if(vulnerabilitySheet.getCell(i, Constant.WavsepTypeColumn.typeTruePositive.getColumn()).getCellType() == Cell.CELL_TYPE_BOOLEAN
                    && vulnerabilitySheet.getCell(i, Constant.WavsepTypeColumn.typeTruePositive.getColumn()).getBooleanCellValue()){
                if(falseNegativeList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.WavsepDetectedColumn.detectedTruePositiveFalse.getColumn());
                    detectingCell.setCellValue(false);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.WavsepDetectedColumn.detectedTruePositiveTrue.getColumn());
                    detectingCell.setCellValue(true);
                }
                // detected false positive type
            } else if (vulnerabilitySheet.getCell(i, Constant.WavsepTypeColumn.typeFalsePositive.getColumn()).getCellType() == Cell.CELL_TYPE_BOOLEAN
                    && !vulnerabilitySheet.getCell(i, Constant.WavsepTypeColumn.typeFalsePositive.getColumn()).getBooleanCellValue()){
                if(trueNegativeList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.WavsepDetectedColumn.detectedFalsePositiveTrue.getColumn());
                    detectingCell.setCellValue(true);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.WavsepDetectedColumn.detectedFalsePositiveFalse.getColumn());
                    detectingCell.setCellValue(false);
                }
                // detected experimental type
            } else if (vulnerabilitySheet.getCell(i, Constant.WavsepTypeColumn.typeExperimental.getColumn()).getCellType() == Cell.CELL_TYPE_STRING
                    && vulnerabilitySheet.getCell(i, Constant.WavsepTypeColumn.typeExperimental.getColumn()).getStringCellValue().equals("EX")){
                if(failedScanningExperimentalList.contains(testCaseName)){
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.WavsepDetectedColumn.detectedExperimentalFalse.getColumn());
                    detectingCell.setCellValue(false);
                } else {
                    detectingCell = vulnerabilitySheet.getCell(i, Constant.WavsepDetectedColumn.detectedExperimentalTrue.getColumn());
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
                if(!wavsepSheet.toString().equals(Constant.WavsepSheets.total.getSheetName())) {
                    writeVulnerabilitySheetContents(wavsepSheet.getSheetName(), FileUtil.readFile(crawledFilePath), FileUtil.readFile(scannedFilePath));
                }
            }
        }
    }

    /**
     * get failed crawling list in a vulnerability sheet
     * @param testCases test cases map
     * @param pureCrawledList pure crawled list, not parsing list
     * @return failed crawling list
     */
    private List<String> getFailedCrawlingList(Map<String, Short> testCases, List<String> pureCrawledList){
        List<String> expectedCrawlingList = new ArrayList<>();
        expectedCrawlingList.addAll(testCases.keySet());
        return WavsepParser.getFailedCrawlingList(pureCrawledList, expectedCrawlingList);
    }

    /**
     * get false negative list ( failed scanning true positive list ) in a vulnerability sheet
     *
     * @param testCases test cases map
     * @param pureScannedList pure scanned list, not parsing list
     * @return false negative list
     */
    private List<String> getFalseNegativeList(Map<String, Short> testCases, List<String> pureScannedList){
        List<String> expectedFalseNegativeList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_TRUE_POSITIVE){
                expectedFalseNegativeList.add(key);
            }
        }
        return WavsepParser.getFalseNegativeList(pureScannedList, expectedFalseNegativeList);
    }

    /**
     * get true negative list ( failed scanning false positive list ) in a vulnerability sheet
     *
     * @param testCases test cases map
     * @param pureScannedList scanned file path
     * @return true negative list
     */
    private List<String> getTrueNegativeList(Map<String, Short> testCases, List<String> pureScannedList){
        List<String> expectedTrueNegativeList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_FALSE_POSITIVE){
                expectedTrueNegativeList.add(key);
            }
        }
        return WavsepParser.getTrueNegativeList(pureScannedList, expectedTrueNegativeList);
    }

    /**
     * get failed scanning experimental list
     *
     * @param testCases test cases map
     * @param pureScannedList pure scanned list, not parsing list
     * @return failed scanning experimental list
     */
    private List<String> getFailedScanningExperimentalList(Map<String, Short> testCases, List<String> pureScannedList){
        List<String> expectedScanningExperimentalList = new ArrayList<>();
        for(String key : testCases.keySet()){
            if(testCases.get(key) == TYPE_EXPERIMENTAL){
                expectedScanningExperimentalList.add(key);
            }
        }
        return WavsepParser.getExperimentalList(pureScannedList, expectedScanningExperimentalList);
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

        while((testCaseCell = vulnerabilitySheet.getCell(testCaseRow, Constant.TEST_CASES_COLUMN)).getCellType() != Cell.CELL_TYPE_BLANK){
            Cell typeCell;
            if((typeCell = vulnerabilitySheet.getCell(testCaseRow, Constant.WavsepTypeColumn.typeTruePositive.getColumn())).getCellType() == Cell.CELL_TYPE_BOOLEAN && typeCell.getBooleanCellValue()){
                testCases.put(testCaseCell.getStringCellValue(), TYPE_TRUE_POSITIVE);
            } else if ((typeCell = vulnerabilitySheet.getCell(testCaseRow, Constant.WavsepTypeColumn.typeFalsePositive.getColumn())).getCellType() == Cell.CELL_TYPE_BOOLEAN && !typeCell.getBooleanCellValue()){
                testCases.put(testCaseCell.getStringCellValue(), TYPE_FALSE_POSITIVE);
            } else if((typeCell = vulnerabilitySheet.getCell(testCaseRow, Constant.WavsepTypeColumn.typeExperimental.getColumn())).getCellType() == Cell.CELL_TYPE_STRING && typeCell.getStringCellValue().equals("EX")){
                testCases.put(testCaseCell.getStringCellValue(), TYPE_EXPERIMENTAL);
            }
            testCaseRow++;
        }

        return testCases;
    }

    /**
     * get wavsep workbook
     *
     * @return wavsep tc workbook
     */
    public TcWorkbook getWavsepWorkbook() {
        return wavsepWorkbook;
    }

    /**
     * get wavsep template file path
     *
     * @return wavsep template file path
     */
    public String getWavsepTemplateFilePath() {
        return wavsepTemplateFilePath;
    }

    /**
     * write excel file
     *
     * @param parentDirectoryPath parent directory path
     * @param fileName file name
     */
    public void writeExcelFile(String parentDirectoryPath, String fileName){
        TcUtil.writeTcWorkbook(parentDirectoryPath, fileName, this.wavsepWorkbook);
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
