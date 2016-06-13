package org.poi.BenchmarkPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.poi.Constant;
import org.poi.POIConstant;
import org.poi.TestCasePOI.TcSheet;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.CellStylesUtil;
import org.poi.Util.FileUtil;
import org.poi.Util.TcUtil;

import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;

/**
 * Created by Hwan on 2016-06-07.
 */
public class BenchmarkCompareTemplate {

    /** before TcWorkbook for compare */
    private TcWorkbook beforeWorkbook = null;
    /** after TcWorkbook for compare */
    private TcWorkbook afterWorkbook = null;
    /** compare TcWorkbook for compare */
    private TcWorkbook compareWorkbook = null;

    /** time generated a benchmark compare template */
    private Date date;

    /**cellStyleUtil instance for cell styles */
    private CellStylesUtil cellStyle;

    /** row number start bnechmark test cases for compare */
    private static final int START_BENCHMARK_TC_ROW = 7;

    /** row nubmer start benchmark compare test cases for compare */
    private static final int START_BENCHMARK_COMPARE_ROW = 6;

    private String beforeFileName = "";
    private String afterFileName = "";


    /**
     * BenchmarkCompareTemplate instant
     *
     * @param beforeWorkbook a before benchmark workbook
     * @param afterWorkbook a after benchmark workbook
     */
    public BenchmarkCompareTemplate(TcWorkbook beforeWorkbook, TcWorkbook afterWorkbook){
        this.beforeWorkbook = beforeWorkbook;
        this.afterWorkbook = afterWorkbook;
        this.compareWorkbook = new TcWorkbook();

        init();
    }

    /**
     * BenchmarkCompareTemplate instant
     *
     * @param beforePath a before benchmark workbook path
     * @param afterPath a after benchmark workbook path
     * @throws IOException
     */
    public BenchmarkCompareTemplate(String beforePath, String afterPath) throws IOException {
        this.beforeWorkbook = new TcWorkbook(FileUtil.readExcelFile(beforePath));
        this.afterWorkbook = new TcWorkbook(FileUtil.readExcelFile(afterPath));
        this.compareWorkbook = new TcWorkbook();

        setFileNames(beforePath, afterPath);

        init();
    }

    /**
     * set file Names
     *
     * @param beforePath before file path
     * @param afterPath after file path
     */
    private void setFileNames(String beforePath, String afterPath){
        this.beforeFileName = new File(beforePath).getName();
        this.afterFileName = new File(afterPath).getName();
    }

    /** initialize Benchmark compare template */
    private void init(){

        this.date = new Date();
        this.cellStyle = new CellStylesUtil(compareWorkbook.getWorkbook());

        for(Constant.BenchmarkSheets sheet : Constant.BenchmarkSheets.values()){
            if(sheet.toString().equals(Constant.BenchmarkSheets.total.toString())){
                setTotalCompareSheet(this.compareWorkbook.getTcSheet(sheet.getSheetName()));
            } else {
                setVulnerabilityCompareSheet(this.compareWorkbook.getTcSheet(sheet.getSheetName()));
            }
        }

        // compare all benchmark vulnerability sheets
        compareVulnerabilitySheets();
        // copy before adn after total sheet in compare total sheet
        compareTotalSheet();
    }

    /**
     * set Total sheet
     * @param totalSheet total sheet
     */
    private void setTotalCompareSheet(TcSheet totalSheet){
        setTotalCompareStyles(totalSheet);
        setTotalCompareWidth(totalSheet);
        setTotalCompareValues(totalSheet);
    }

    /**
     * set compare total sheet styles
     *
     * @param sheet total sheet
     */
    private void setTotalCompareStyles(TcSheet sheet){
        sheet.setSameCellStyle(1, 1, 'a', 'f', cellStyle.getSimpleCellStyle(true, false, 0, 2, 0, 0));
        sheet.setSameCellStyle(2, 2, 'b', 'e', cellStyle.getSimpleCellStyle(true, true, 2, 2, 1, 1));
        sheet.setSameCellStyle(3, 13, 'b', 'e', cellStyle.getSimpleCellStyle(false, false, 1, 1, 1, 1));
        sheet.setSameCellStyle(3, 13, 'a', 'a', cellStyle.getSimpleCellStyle(false, false, 1, 1, 1, 2));
        sheet.setSameCellStyle(14, 14, 'b', 'e', cellStyle.getSimpleCellStyle(false, false, 2, 2, 1, 1));
        sheet.setSameCellStyle(2, 14, 'g', 'g', cellStyle.getSimpleCellStyle(false, false, 0, 0, 2, 0));
        sheet.setSameCellStyle(3, 13, 'f', 'f', cellStyle.getSimpleCellStyle(false, false, 1, 1, 2, 2));
        sheet.setSameCellStyle(15, 15, 'a', 'f', cellStyle.getSimpleCellStyle(false, false, 2, 0, 0, 0));
        sheet.getCell(2, 'a').setCellStyle(cellStyle.getSimpleCellStyle(true, true, 2, 2, 2, 2));
        sheet.getCell(14, 'a').setCellStyle(cellStyle.getSimpleCellStyle(true, false, 2, 2, 2, 2));
        sheet.getCell(2, 'f').setCellStyle(cellStyle.getSimpleCellStyle(true, true, 2, 2, 2, 2));
        sheet.getCell(14, 'f').setCellStyle(cellStyle.getSimpleCellStyle(true, false, 2, 2, 2, 2));
        sheet.getCell(1, 'b').getCellStyle().setDataFormat(compareWorkbook.getWorkbook().getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd HH:mm"));
    }


    /**
     * set compare total sheet width
     *
     * @param sheet total sheet
     */
    private void setTotalCompareWidth(TcSheet sheet){
        // set a total compare sheet default width column 'A'
        sheet.setColumnWidth('a', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);

        // set total compare sheet default width column from 'B' to 'I'
        for(char i = 'b'; i <= 'j'; i++){
            sheet.setColumnWidth(i, POIConstant.COMPARE_DEFAULT_COLUMN_WIDTH);
        }
    }

    /**
     * set compare total sheet values
     *
     * @param sheet total sheet
     */
    private void setTotalCompareValues(TcSheet sheet){
        sheet.getCell(1, 'a').setCellValue("Compare OWASP Benchmark");
        sheet.getCell(1, 'b').setCellValue(this.date);
        sheet.getCell(2, 'a').setCellValue("Category");
        sheet.getCell(2, 'b').setCellValue("Both True");
        sheet.getCell(2, 'c').setCellValue("Both False");
        sheet.getCell(2, 'd').setCellValue("Only Before True");
        sheet.getCell(2, 'e').setCellValue("Only After True");
        sheet.getCell(2, 'f').setCellValue("Total");

        //todo modify values
        int row = 3;
        for (Constant.BenchmarkSheets benchmarkSheet : Constant.BenchmarkSheets.values()) {
            if (!benchmarkSheet.toString().equals(Constant.BenchmarkSheets.total.toString())) {
                sheet.getCell(row, 'a').setCellValue(benchmarkSheet.getSheetName());
                sheet.getCell(row, 'b').setCellFormula("\'" + benchmarkSheet.getSheetName() + "\'" + "!C5+" + "\'" + benchmarkSheet.getSheetName() + "\'" + "!G5");
                sheet.getCell(row, 'c').setCellFormula("\'" + benchmarkSheet.getSheetName() + "\'" + "!D5+" + "\'" + benchmarkSheet.getSheetName() + "\'" + "!H5");
                sheet.getCell(row, 'd').setCellFormula("\'" + benchmarkSheet.getSheetName() + "\'" + "!E5+" + "\'" + benchmarkSheet.getSheetName() + "\'" + "!I5");
                sheet.getCell(row, 'e').setCellFormula("\'" + benchmarkSheet.getSheetName() + "\'" + "!F5+" + "\'" + benchmarkSheet.getSheetName() + "\'" + "!J5");
                sheet.getCell(row, 'f').setCellFormula("SUM(B" + row + ":E" + row + ")");

                row++;
            }
        }

        sheet.getCell(row, 'a').setCellValue("Totals");
        sheet.getCell(row, 'b').setCellFormula("SUM(B" + 4 + ":B" + (row - 1) + ")");
        sheet.getCell(row, 'c').setCellFormula("SUM(C" + 4 + ":C" + (row - 1) + ")");
        sheet.getCell(row, 'd').setCellFormula("SUM(D" + 4 + ":D" + (row - 1) + ")");
        sheet.getCell(row, 'e').setCellFormula("SUM(E" + 4 + ":E" + (row - 1) + ")");
        sheet.getCell(row, 'f').setCellFormula("SUM(F" + 4 + ":F" + (row - 1) + ")");

    }

    /**
     * add before and after total sheet content in compare total sheet
     */
    private void compareTotalSheet(){
        String totalSheetName = Constant.WavsepSheets.total.getSheetName();
        copyBeforeAndAfterTotalInCompareTotal(compareWorkbook.getTcSheet(totalSheetName), beforeWorkbook.getTcSheet(totalSheetName), afterWorkbook.getTcSheet(totalSheetName));
    }

    /**
     * copy before and after total sheet contents in compare total sheet
     *
     * @param compareTotal compare total sheet
     * @param beforeTotal before total sheet
     * @param afterTotal after total sheet
     */
    private void copyBeforeAndAfterTotalInCompareTotal(TcSheet compareTotal, TcSheet beforeTotal, TcSheet afterTotal){
        int compareRow = 19;

        compareTotal.getCell(compareRow++, 'a').setCellValue("Before : " + beforeFileName);
        compareRow = copySheet(compareTotal, beforeTotal, compareRow, 0) + 2;
        compareTotal.getCell(compareRow++, 'a').setCellValue("After : " + afterFileName);

        copySheet(compareTotal, afterTotal, compareRow, 0);
    }

    /**
     * copied sheet is copied to the target sheet.
     *
     * @param target taget sheet
     * @param copied copied sheet
     * @param startRow start row number
     * @param startColumn start column number
     * @return next row number
     */
    private int copySheet(TcSheet target, TcSheet copied, int startRow, int startColumn){
        Iterator<Row> copiedRowIterator = copied.getSheet().rowIterator();

        setBeforeAndAfterTotalStyle(target, startRow);

        while(copiedRowIterator.hasNext()){
            Row beforeRow = copiedRowIterator.next();
            Iterator<Cell> copiedCellIterator = beforeRow.cellIterator();
            int copyColumn = startColumn;
            while(copiedCellIterator.hasNext()){
                Cell copiedCell = copiedCellIterator.next();
                Cell compareCell = target.getCell(startRow - 1, copyColumn);
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

                copyColumn++;
            }
            startRow++;
        }

        return startRow;
    }

    /**
     * set before and after total style
     *
     * @param target target sheet
     * @param startRow start row to copy
     */
    private void setBeforeAndAfterTotalStyle(TcSheet target, int startRow){
        int tempRow = startRow;
        int vulnerabilityRowHeight = 9;
        int totalRowHeight = 1;

        // percent cells style (2, h) ~ (11, j), (13, h) ~ (15, j)
        CellStyle percentCellStyle = cellStyle.getSimpleCellStyle(false, false, 1, 1, 1, 1);
        percentCellStyle.setDataFormat(compareWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        // percent bottom cells style (12, h) ~ (12, j)
        CellStyle percentBottomCellStyle = cellStyle.getSimpleCellStyle(false, false, 1, 2, 1, 1);
        percentBottomCellStyle.setDataFormat(compareWorkbook.getWorkbook().createDataFormat().getFormat("0%"));

        // set top border
        target.setSameCellStyle(tempRow - 1, tempRow - 1, 'a', 'j', cellStyle.getSimpleCellStyle(true, false, 0, 2, 0, 0));

        // set category cells style (1, 'a') ~ (1, 'j')
        target.getCell(tempRow, 'a').setCellStyle(cellStyle.getSimpleCellStyle(true, true, 2, 2, 0, 2));
        target.setSameCellStyle(tempRow, tempRow, 'b', 'f', cellStyle.getSimpleCellStyle(true, true, 2, 2, 0, 0));
        target.getCell(tempRow, 'g').setCellStyle(cellStyle.getSimpleCellStyle(true, true, 2, 2, 0, 2));
        target.setSameCellStyle(tempRow, tempRow, 'h', 'j', cellStyle.getSimpleCellStyle(true, true, 2, 2, 0 ,0));

        // set vulnerability cells style (2, 'a') ~ (11, j)
        target.setSameCellStyle(++tempRow, tempRow + vulnerabilityRowHeight, 'a', 'a', cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 2));
        target.setSameCellStyle(tempRow, tempRow + vulnerabilityRowHeight, 'b', 'f', cellStyle.getSimpleCellStyle(false, false, 1, 1, 0 ,0));
        target.setSameCellStyle(tempRow, tempRow + vulnerabilityRowHeight, 'g', 'g', cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 2));
        target.setSameCellStyle(tempRow, tempRow + vulnerabilityRowHeight, 'h', 'j', percentCellStyle);
        tempRow = tempRow + vulnerabilityRowHeight;

        // set vulnerability bottom cells style (12, 'a') ~ (12, 'j')
        target.getCell(++tempRow, 'a').setCellStyle(cellStyle.getSimpleCellStyle(false, false, 1, 2, 0, 2));
        target.setSameCellStyle(tempRow, tempRow, 'b', 'f' , cellStyle.getSimpleCellStyle(false, false, 1, 2, 0, 0));
        target.getCell(tempRow, 'g').setCellStyle(cellStyle.getSimpleCellStyle(false, false, 1, 2, 0, 2));
        target.setSameCellStyle(tempRow, tempRow, 'h', 'i', percentBottomCellStyle);

        // set total cells style (13, 'a') ~ (14, 'j')
        target.setSameCellStyle(++tempRow, tempRow + totalRowHeight, 'a', 'a', cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 2));
        target.setSameCellStyle(tempRow, tempRow + totalRowHeight, 'b', 'f', cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 0));
        target.setSameCellStyle(tempRow, tempRow + totalRowHeight, 'g', 'g', cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 2));
        target.setSameCellStyle(tempRow, tempRow + totalRowHeight, 'h' , 'j', percentCellStyle);
        tempRow = tempRow + totalRowHeight;

        // set right border
        target.setSameCellStyle(startRow, tempRow, 'k' ,'k', cellStyle.getSimpleCellStyle(false, false, 0, 0, 2, 0));
        // set left border
        target.setSameCellStyle(++tempRow, tempRow, 'a', 'j', cellStyle.getSimpleCellStyle(false, false, 2, 0, 0, 0));
    }

    /**
     * set a vulnerability sheet
     *
     * @param vulnerabilitySheet vulnerability sheet
     */
    private void setVulnerabilityCompareSheet(TcSheet vulnerabilitySheet){
        setVulnerabilityCompareStyles(vulnerabilitySheet);
        setVulnerabilityCompareWidth(vulnerabilitySheet);
        setVulnerabilityCompareValues(vulnerabilitySheet);
    }

    /**
     * set a compare vulnerability sheet styles
     *
     * @param sheet vulnerability sheet
     */
    private void setVulnerabilityCompareStyles(TcSheet sheet){
        sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex('a'), cellStyle.DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);

        for(int i = TcUtil.convertColumnAlphabetToIndex('c'); i <= TcUtil.convertColumnAlphabetToIndex('j'); i++) {
            sheet.getSheet().setDefaultColumnStyle(i, cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        sheet.setSameCellStyle(1, 1, 'a', 'j', cellStyle.getSimpleCellStyle(true, false, 0, 2, 0, 0));
        sheet.setSameCellStyle(2, 4, 'b', 'j', cellStyle.getSimpleCellStyle(true, true, 1, 1, 1, 1));
        sheet.setSameCellStyle(5, 5, 'a', 'j', cellStyle.getSimpleCellStyle(true, true, 1, 2, 1, 1));
    }

    /**
     * set a compare vulnerability sheet width
     *
     * @param sheet vulnerability sheet
     */
    private void setVulnerabilityCompareWidth(TcSheet sheet){
        sheet.setColumnWidth('a', POIConstant.DEFAULT_COLUMN_WIDTH);
        sheet.setColumnWidth('b', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        for(char i = 'c'; i <= 'j'; i++){
            sheet.setColumnWidth(i, POIConstant.COMPARE_DEFAULT_COLUMN_WIDTH);
        }
    }


    /**
     * set a compare vulnerability sheet values
     *
     * @param sheet vulnerability sheet
     */
    private void setVulnerabilityCompareValues(TcSheet sheet){
        sheet.getCell(1, 'a').setCellValue(sheet.getSheetName());
        sheet.getCell(5, 'a').setCellValue("합 계");
        sheet.mergeCellRegion(2, 4, 'b', 'b').setCellValue("Test Cases");
        sheet.mergeCellRegion(2, 2, 'c', 'j').setCellValue("Compare");
        sheet.mergeCellRegion(3, 3, 'c', 'f').setCellValue("True Positive");
        sheet.mergeCellRegion(3, 3, 'g', 'j').setCellValue("False Positive");

        sheet.getCell(4, 'c').setCellValue("Both True");
        sheet.getCell(4, 'd').setCellValue("Both False");
        sheet.getCell(4, 'e').setCellValue("Only Before True");
        sheet.getCell(4, 'f').setCellValue("Only After True");
        sheet.getCell(4, 'g').setCellValue("Both True");
        sheet.getCell(4, 'h').setCellValue("Both False");
        sheet.getCell(4, 'i').setCellValue("Only Before True");
        sheet.getCell(4, 'j').setCellValue("Only After True");

        sheet.getCell(5, 'b').setCellFormula("COUNTA(B:B)-2");
        sheet.getCell(5, 'c').setCellFormula("COUNTIF(C:C,\"TRUE\")");
        sheet.getCell(5, 'd').setCellFormula("COUNTIF(D:D, \"FALSE\")");
        sheet.getCell(5, 'e').setCellFormula("COUNTIF(E:E, \"TRUE\")");
        sheet.getCell(5, 'f').setCellFormula("COUNTIF(F:F, \"TRUE\")");
        sheet.getCell(5, 'g').setCellFormula("COUNTIF(G:G,\"TRUE\")");
        sheet.getCell(5, 'h').setCellFormula("COUNTIF(H:H, \"FALSE\")");
        sheet.getCell(5, 'i').setCellFormula("COUNTIF(I:I, \"TRUE\")");
        sheet.getCell(5, 'j').setCellFormula("COUNTIF(J:J, \"TRUE\")");
    }

    /**
     * compare vulnerability sheets
     */
    private void compareVulnerabilitySheets(){
        for(Constant.BenchmarkSheets benchmarkSheet : Constant.BenchmarkSheets.values()){
            if(!benchmarkSheet.toString().equals(Constant.BenchmarkSheets.total.toString())){
                String sheetName = benchmarkSheet.getSheetName();
                compareBeforeAndAfterVulnerabilitySheet(compareWorkbook.getTcSheet(sheetName), beforeWorkbook.getTcSheet(sheetName), afterWorkbook.getTcSheet(sheetName));
            }
        }
    }

    /**
     * write compare before sheet and after sheet in compare sheet
     *
     * @param compareSheet the compare sheet
     * @param beforeSheet the before sheet
     * @param afterSheet the after sheet
     */
    private void compareBeforeAndAfterVulnerabilitySheet(TcSheet compareSheet, TcSheet beforeSheet, TcSheet afterSheet){
        int tcRow = START_BENCHMARK_TC_ROW;
        int compareRow = START_BENCHMARK_COMPARE_ROW;

        while(afterSheet.getCell(tcRow, 'b').getCellType() != Cell.CELL_TYPE_BLANK) {
            try {
                /** after sheet test case cell string value */
                String afterCellTC = afterSheet.getCell(tcRow, 'b').getStringCellValue();
                /** before sheet test case cell string value */
                String beforeCellTC = beforeSheet.getCell(tcRow, 'b').getStringCellValue();
                /** compare sheet test case cell */
                Cell compareCell = compareSheet.getCell(compareRow, 'b');
                if (afterCellTC.equals(beforeCellTC)) {
                    // write test case in Compare sheet
                    compareCell.setCellValue(afterCellTC);
                    compareCell.setCellStyle(cellStyle.DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);

                    if (!(afterSheet.getCell(tcRow, 'c').getCellType() == Cell.CELL_TYPE_BLANK) && afterSheet.getCell(tcRow, 'c').getBooleanCellValue()) {
                        /** tc is true positive */
                        if (getCellValue(afterSheet.getCell(tcRow, 'g')) != null && getCellValue(afterSheet.getCell(tcRow, 'g')).toString().equals("TRUE")) {
                            if (getCellValue(beforeSheet.getCell(tcRow, 'g')) != null && getCellValue(beforeSheet.getCell(tcRow, 'g')).toString().equals("TRUE")) {
                                /** both true */
                                compareSheet.getCell(compareRow, 'c').setCellValue(true);
                            } else {
                                /** only after true */
                                // beforeSheet.getCell(tcRow, 'h').getStringCellValue().equals("FALSE") same meaning
                                compareSheet.getCell(compareRow, 'f').setCellValue(true);
                            }
                        } else {
                            // afterSheet.getCell(tcRow, 'h').getStringCellValue().equals("FALSE") same meaning
                            if (getCellValue(beforeSheet.getCell(tcRow, 'h')) != null && getCellValue(beforeSheet.getCell(tcRow, 'h')).toString().equals("FALSE")) {
                                /** both false */
                                compareSheet.getCell(compareRow, 'd').setCellValue(false);
                            } else {
                                /** only before true */
                                // before.getCell(tcRow, 'g').getStringCellValue().equals("TRUE") same meaning
                                compareSheet.getCell(compareRow, 'e').setCellValue(true);
                            }
                        }
                    } else if (!(afterSheet.getCell(tcRow, 'd').getCellType() == Cell.CELL_TYPE_BLANK) && !afterSheet.getCell(tcRow, 'd').getBooleanCellValue()) {
                        /** tc is false positive */
                        if (getCellValue(afterSheet.getCell(tcRow, 'i')) != null && getCellValue(afterSheet.getCell(tcRow, 'i')).toString().equals("TRUE")) {
                            if (getCellValue(beforeSheet.getCell(tcRow, 'i')) != null && getCellValue(beforeSheet.getCell(tcRow, 'i')).toString().equals("TRUE")) {
                                /** both true */
                                compareSheet.getCell(compareRow, 'g').setCellValue(true);
                            } else {
                                /** only after true */
                                // beforeSheet.getCell(tcRow, 'j').getStringCellValue().equals("FALSE") same meaning
                                compareSheet.getCell(compareRow, 'j').setCellValue(true);
                            }
                        } else {
                            // afterSheet.getCell(tcRow, 'j').getStringCellValue().equals("FALSE") same meaning
                            if (getCellValue(beforeSheet.getCell(tcRow, 'j')) != null && getCellValue(beforeSheet.getCell(tcRow, 'j')).toString().equals("FALSE")) {
                                /** both false */
                                compareSheet.getCell(compareRow, 'h').setCellValue(false);
                            } else {
                                /** only before true */
                                // before.getCell(tcRow, 'i').getStringCellValue().equals("TRUE") same meaning
                                compareSheet.getCell(compareRow, 'i').setCellValue(true);
                            }
                        }
                    }
                } else {
                    /** after test case and before case is not same */
                    compareCell.setCellStyle(cellStyle.DEFAULT_THICK_RIGHT_LEFT_BG_GRAY);
                }
            } catch (Exception e){
                e.printStackTrace();
            } finally {
                // move next row
                tcRow++;
                compareRow++;
            }
        }
    }

    /**
     * get a cell value
     *
     * @param cell a cell
     */
    private Object getCellValue(Cell cell){
        int cellType;
        if(cell.getCellType() != Cell.CELL_TYPE_FORMULA){
            cellType = cell.getCellType();
        } else {
            cellType = cell.getCachedFormulaResultType();
        }

        switch(cellType){
            case Cell.CELL_TYPE_BOOLEAN :
                return cell.getBooleanCellValue();
            case Cell.CELL_TYPE_NUMERIC :
                return cell.getNumericCellValue();
            case Cell.CELL_TYPE_STRING :
                return cell.getStringCellValue();
            case Cell.CELL_TYPE_ERROR :
                return cell.getErrorCellValue();
            default :
                return null;
        }
    }


    /**
     * Write current workbook state in a file
     *
     * @param outputDir output directory path
     * @param fileName file name
     * @return file absolute path
     */
    public String writeCompareWorkbook(String outputDir, String fileName){
        if(compareWorkbook != null){
            return compareWorkbook.writeWorkbook(outputDir, fileName);
        }
        return null;
    }

    public static void main(String[] args){
        try {
            BenchmarkCompareTemplate compare = new BenchmarkCompareTemplate("C:\\scalaProjects\\testPOI\\Zap_Benchmark_Results160609153221.xlsx",
                    "C:\\scalaProjects\\testPOI\\Zap_Benchmark_Results160609165623.xlsx");

            compare.writeCompareWorkbook("./", "testBenchmark.xlsx");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}