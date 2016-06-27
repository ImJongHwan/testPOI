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
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;

/**
 * Created by Hwan on 2016-06-07.
 */
public class WavsepComparer {

    private String wavsepCompareTemplateFilePath;
    private TcWorkbook wavsepCompareWorkbook;

    private static final int VULNERABILITY_CONTENTS_START_ROW = 7;
    private static final int COMPARE_VULNERABILITY_START_ROW = 6;
    private static final String TEST_CASES_COLUMN = Constant.TEST_CASES_COLUMN;

    private static final String DEFAULT_WAVSEP_COMPARE_TEMPLATE_RESOURCE_FILE_PATH = "Template/wavsepCompareTemplate.xlsx";

    public WavsepComparer(String wavsepCompareTemplateFilePath) throws IOException {
        this.wavsepCompareTemplateFilePath = wavsepCompareTemplateFilePath;
        this.wavsepCompareWorkbook = new TcWorkbook(FileUtil.readExcelFile(wavsepCompareTemplateFilePath));
    }

    public WavsepComparer() throws IOException{
        this.wavsepCompareTemplateFilePath = DEFAULT_WAVSEP_COMPARE_TEMPLATE_RESOURCE_FILE_PATH;
        XSSFWorkbook readWorkbook = new XSSFWorkbook(WavsepComparer.class.getClassLoader().getResourceAsStream(DEFAULT_WAVSEP_COMPARE_TEMPLATE_RESOURCE_FILE_PATH));
        this.wavsepCompareWorkbook = new TcWorkbook(readWorkbook);
    }

    /**
     * write compare contents in a benchmark compare workbook
     *
     * @param beforeWorkbookPath before workbook file path
     * @param afterWorkbookPath after workbook file path
     * @throws IOException
     */
    public void writeCompareInWorkbook(String beforeWorkbookPath, String afterWorkbookPath) throws IOException {
        TcWorkbook beforeWavsepWorkbook = new TcWorkbook(FileUtil.readExcelFile(beforeWorkbookPath));
        TcWorkbook afterWavsepWorkbook = new TcWorkbook(FileUtil.readExcelFile(afterWorkbookPath));

        for(Constant.WavsepSheets wavsepSheet : Constant.WavsepSheets.values()){
            String wavsepSheetName = wavsepSheet.getSheetName();
            if(wavsepSheet.toString().equals("total")){
                setBeforeAndAfterWorkbookTitleInTotalSheet(wavsepCompareWorkbook.getTcSheet(wavsepSheetName), beforeWorkbookPath, afterWorkbookPath);
                writeExistingTotalsInTotalSheet(wavsepCompareWorkbook.getTcSheet(wavsepSheetName), beforeWavsepWorkbook.getTcSheet(wavsepSheetName), afterWavsepWorkbook.getTcSheet(wavsepSheetName));
            } else {
                writeCompareInVulnerabilitySheet(wavsepCompareWorkbook.getTcSheet(wavsepSheetName), beforeWavsepWorkbook.getTcSheet(wavsepSheetName), afterWavsepWorkbook.getTcSheet(wavsepSheetName));
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

        compareTotalSheet.getCell(BEFORE_TITLE_ROW, BEFORE_TITLE_COLUMN).setCellValue(getFileNameFromPath(beforeWorkbookPath) + " (Before)");
        compareTotalSheet.getCell(AFTER_TITLE_ROW, AFTER_TITLE_COLUMN).setCellValue(getFileNameFromPath(afterWorkbookPath) + " (After)");
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
        final int startRowForCopyBefore = 18;
        final int startRowForCopyAfter = 32;

        copyTotalContents(beforeTotalSheet, compareTotalSheet, startRowForCopyBefore);
        copyTotalContents(afterTotalSheet, compareTotalSheet, startRowForCopyAfter);
    }

    /**
     * copy total contents in a compare total sheet
     *
     * @param copiedWavsepTotalSheet copied wavsep total sheet
     * @param compareWavsepTotalSheet compare wavsep total sheet
     * @param compareTotalRowForCopy compare total row for copy
     */
    private void copyTotalContents(TcSheet copiedWavsepTotalSheet, TcSheet compareWavsepTotalSheet, int compareTotalRowForCopy){
        final int copiedStartRow = 1;
        final int copiedEndRow = 12;

        // copy rows
        for(int i = copiedStartRow; i <= copiedEndRow; i++){
            Iterator<Cell> cellIterator = copiedWavsepTotalSheet.getSheet().getRow(i -1).iterator();
            while(cellIterator.hasNext()){
                Cell copiedCell = cellIterator.next();
                Cell compareCell = compareWavsepTotalSheet.getCell(compareTotalRowForCopy - 1, copiedCell.getColumnIndex());

                // copy cell value
                int copiedCellType = copiedCell.getCellType();
                if(copiedCellType == Cell.CELL_TYPE_FORMULA){
                    copiedCellType = copiedCell.getCachedFormulaResultType();
                }

                switch(copiedCellType){
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

                // copy cell style
//                compareCell.getCellStyle().cloneStyleFrom(copiedCell.getCellStyle());
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
        // TODO: 2016-06-23
        Map<String, Boolean> beforeCrawlingMap = readCrawlingTestCases(beforeVulnerabilitySheet);
        Map<String, Boolean> afterCrawlingMap = readCrawlingTestCases(afterVulnerabilitySheet);
        Map<String, Boolean> beforeDetectingMap = readDetectingTestCases(beforeVulnerabilitySheet);
        Map<String, Boolean> afterDetectingMap = readDetectingTestCases(afterVulnerabilitySheet);

        Cell testCaseCell;

        for(int i = COMPARE_VULNERABILITY_START_ROW; (testCaseCell = compareVulnerabilitySheet.getCell(i, TEST_CASES_COLUMN)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            String testCaseName = testCaseCell.getStringCellValue();
            CellStylesUtil cellStyle = new CellStylesUtil(this.wavsepCompareWorkbook.getWorkbook());
            CellStyle simpleStyle = cellStyle.getSimpleCellStyle(false, true, 0, 0, 1, 1);
            Cell compareCell;
            short flagDiff = 0;

            if(beforeCrawlingMap.get(testCaseName)){
                compareCell  = compareVulnerabilitySheet.getCell(i, "F");
                compareCell.setCellValue(true);
                compareCell.setCellStyle(simpleStyle);
            }

            if(afterCrawlingMap.get(testCaseName)){
                compareCell  = compareVulnerabilitySheet.getCell(i, "G");
                compareCell.setCellValue(true);
                compareCell.setCellStyle(simpleStyle);
            }

            if(beforeDetectingMap.get(testCaseName)){
                compareCell  = compareVulnerabilitySheet.getCell(i, "H");
                compareCell.setCellValue(true);
                compareCell.setCellStyle(simpleStyle);
                flagDiff++;
            }

            if(afterDetectingMap.get(testCaseName)){
                compareCell  = compareVulnerabilitySheet.getCell(i, "I");
                compareCell.setCellValue(true);
                compareCell.setCellStyle(simpleStyle);
                flagDiff++;
            }

            if(flagDiff == 1){
                if(beforeDetectingMap.get(testCaseName)){
                    compareCell = compareVulnerabilitySheet.getCell(i, "J");
                } else {
                    compareCell = compareVulnerabilitySheet.getCell(i, "K");
                }
                compareCell.setCellValue(true);
                compareCell.setCellStyle(simpleStyle);
            }
        }
    }

    /**
     * read test cases included crawling result in wavsepvulnerability sheet
     *
     * @param wavsepVulnerabilitySheet wavsep vulnerability sheet
     * @return crawling test cases map, key - test case name, value - isCrawled boolean
     */
    private Map<String, Boolean> readCrawlingTestCases(TcSheet wavsepVulnerabilitySheet){
        Map<String, Boolean> crawlingTestCases = new HashMap<>();
        Cell testCaseCell;

        for(int i = VULNERABILITY_CONTENTS_START_ROW ; (testCaseCell = wavsepVulnerabilitySheet.getCell(i, TEST_CASES_COLUMN)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            Cell trueCrawlCell = wavsepVulnerabilitySheet.getCell(i, Constant.WavsepCrawlColumn.urlCrawlTrue.getColumn());
            Cell falseCrawlCell = wavsepVulnerabilitySheet.getCell(i, Constant.WavsepCrawlColumn.urlCrawlFalse.getColumn());

            if(TcUtil.getCellTypeUnwrappedFormula(trueCrawlCell) == Cell.CELL_TYPE_STRING && trueCrawlCell.getStringCellValue().equals("TRUE")){
                crawlingTestCases.put(testCaseCell.getStringCellValue(), Boolean.valueOf(trueCrawlCell.getStringCellValue()));
            } else if(TcUtil.getCellTypeUnwrappedFormula(falseCrawlCell) == Cell.CELL_TYPE_STRING && falseCrawlCell.getStringCellValue().equals("FALSE")){
                crawlingTestCases.put(testCaseCell.getStringCellValue(), Boolean.valueOf(falseCrawlCell.getStringCellValue()));
            }
        }
        return crawlingTestCases;
    }

    /**
     * read test cases included detecting result in wavsep vulnerability sheet
     *
     * @param wavsepVulnerabilitySheet wavsep vulnerability sheet
     * @return detecting test cases map, key - test case name, value - isDetected boolean
     */
    private Map<String, Boolean> readDetectingTestCases(TcSheet wavsepVulnerabilitySheet) {
        Map<String, Boolean> detectingTestCases = new HashMap<>();
        // TODO: 2016-06-23
        Cell testCaseCell;
        for(int i = VULNERABILITY_CONTENTS_START_ROW; (testCaseCell = wavsepVulnerabilitySheet.getCell(i, TEST_CASES_COLUMN)).getCellType() == Cell.CELL_TYPE_STRING; i++){
            for(Constant.WavsepDetectedColumn wavsepDetected : Constant.WavsepDetectedColumn.values()){
                Cell detectedCell = wavsepVulnerabilitySheet.getCell(i, wavsepDetected.getColumn());
                if(TcUtil.getCellTypeUnwrappedFormula(detectedCell) == Cell.CELL_TYPE_STRING){
                    detectingTestCases.put(testCaseCell.getStringCellValue(), Boolean.valueOf(detectedCell.getStringCellValue()));
                    break;
                }
            }
        }
        return detectingTestCases;
    }

    /**
     * get wavsep compare template file path
     * @return wavsep compare template file path
     */
    public String getWavsepCompareTemplateFilePath() {
        return wavsepCompareTemplateFilePath;
    }

    /**
     * get wavsep compare workbook
     * @return wavsep compare workbook
     */
    public TcWorkbook getWavsepCompareWorkbook() {
        return wavsepCompareWorkbook;
    }

    /**
     * write wavsep compare workbook current state in excel file
     *
     * @param parentDirectoryPath parent directory path
     * @param fileName file name
     */
    public void writeWavsepCompareWorkbookExcelFile(String parentDirectoryPath, String fileName){
        this.wavsepCompareWorkbook.writeWorkbook(parentDirectoryPath, fileName);
    }

    public static void main(String[] args){
        try {
            WavsepComparer wavsepComparer = new WavsepComparer("C:\\scalaProjects\\testPOI\\wavsepCompareTemplate.xlsx");
            wavsepComparer.writeCompareInWorkbook("C:\\scalaProjects\\testPOI\\Zap_Wavsep_Results_160602140051.xlsx", "C:\\scalaProjects\\testPOI\\Zap_Wavsep_Results_160608161934.xlsx");
            wavsepComparer.writeWavsepCompareWorkbookExcelFile("C:\\Users\\Hwan\\Desktop", "test.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
