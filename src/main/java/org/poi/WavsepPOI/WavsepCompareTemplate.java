package org.poi.WavsepPOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFName;
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
public class WavsepCompareTemplate {

    /**
     * time generated a wavsep compare template
     */
    private Date date;

    /**
     * cellStyleUtil instance for cell styles
     */
    private CellStylesUtil cellStyle;

    /**
     * row number start wavsep test cases for compare
     */
    private static final int START_WAVSEP_TC_ROW = 7;
    /**
     * row number start wavsep compare test cases for compare
     */
    private static final int START_WAVSEP_COMPARE_ROW = 6;

    /**
     * before TcWorkbook for compare
     */
    private TcWorkbook beforeWorkbook = null;
    /**
     * after TcWorkbook for compare
     */
    private TcWorkbook afterWorkbook = null;
    /**
     * compare TcWrokbook for compare
     */
    private TcWorkbook compareWorkbook = null;

    private String beforeFileName = "";
    private String afterFileName = "";

    /**
     * generate a WavsepCompareTemplate instance
     *
     * @param beforeWorkbook a before wavsep workbook
     * @param afterWorkbook  a after wavsep workbook
     */
    public WavsepCompareTemplate(TcWorkbook beforeWorkbook, TcWorkbook afterWorkbook) {
        this.beforeWorkbook = beforeWorkbook;
        this.afterWorkbook = afterWorkbook;
        this.compareWorkbook = new TcWorkbook();

        init();
    }

    /**
     * generate a WavsepCompateTemplate instance
     *
     * @param beforePath a before wavsep workbook path
     * @param afterPath  a after wavsep wrokbook path
     * @throws IOException
     */
    public WavsepCompareTemplate(String beforePath, String afterPath) throws IOException {
        this.beforeWorkbook = new TcWorkbook(FileUtil.readExcelFile(beforePath));
        this.afterWorkbook = new TcWorkbook(FileUtil.readExcelFile(afterPath));
        this.compareWorkbook = new TcWorkbook();

        setFileNames(beforePath, afterPath);

        init();
    }

    /**
     * set File Names
     *
     * @param beforePath before file path
     * @param afterPath after file path
     */
    private void setFileNames(String beforePath, String afterPath){
        this.beforeFileName = new File(beforePath).getName();
        this.afterFileName = new File(afterPath).getName();
    }

    /**
     * initialize Wavsep compare template
     */
    private void init() {

        this.date = new Date();
        this.cellStyle = new CellStylesUtil(this.compareWorkbook.getWorkbook());

        for (Constant.WavsepSheets wavsepSheet : Constant.WavsepSheets.values()) {
            if (wavsepSheet.toString().equals(Constant.WavsepSheets.total.toString())) {
                setTotalCompareSheet(this.compareWorkbook.getTcSheet(wavsepSheet.getSheetName()));
            } else {
                setVulnerabilityCompareSheet(this.compareWorkbook.getTcSheet(wavsepSheet.getSheetName()));
            }
        }

        // compare all wavsep vulnerability sheets
        compareVulnerabilitySheets();
        // add total sheet, before and after
        compareTotalSheet();
    }

//    /**
//     * set a Generated Time
//     */
//    private void setGeneratedTime(){
//        SimpleDateFormat sdf = new SimpleDateFormat("yyMMddHHmmss");
//        this.generatedTime = sdf.format(new Date());
//    }

    /**
     * set a wavsep total compare sheet
     *
     * @param sheet total compare sheet
     */
    private void setTotalCompareSheet(TcSheet sheet) {
        setTotalCompareStyles(sheet);
        setTotalCompareWidth(sheet);
        setTotalCompareValues(sheet);
    }

    /**
     * set a wavsep total compare sheet styles
     *
     * @param sheet total compare sheet
     */
    private void setTotalCompareStyles(TcSheet sheet) {
        sheet.setSameCellStyle(4, 8, 'a', 'a', cellStyle.getSimpleCellStyle(false, false, 1, 1, 0, 2));
        sheet.setSameCellStyle(10, 11, 'a', 'a', cellStyle.getSimpleCellStyle(false, false, 1, 1, 0 , 2));
        sheet.setSameCellStyle(4, 11, 'b', 'e', cellStyle.getSimpleCellStyle(false, false, 1, 1, 1, 1));
        sheet.setSameCellStyle(12, 12, 'b', 'e', cellStyle.getSimpleCellStyle(false, false, 2, 2, 1, 1));
        sheet.getCell(12, 'a').setCellStyle(cellStyle.getSimpleCellStyle(true, false, 2, 2, 2, 2));
        sheet.getCell(12, 'f').setCellStyle(cellStyle.getSimpleCellStyle(true, false, 2, 2, 2, 2));
        sheet.getCell(2, 'f').setCellStyle(cellStyle.getSimpleCellStyle(true, true, 2, 2, 2, 2));
        sheet.setSameCellStyle(4, 11, 'f', 'f', cellStyle.getSimpleCellStyle(false, false, 1, 1, 2, 2));
        sheet.setSameCellStyle(1, 1, 'a', 'f', cellStyle.getSimpleCellStyle(true , false, 0, 2, 0, 0));
        sheet.getCell(1, 'b').getCellStyle().setDataFormat(this.compareWorkbook.getWorkbook().getCreationHelper().createDataFormat().getFormat("yyyy-MM-dd HH:mm"));
        sheet.getCell(2, 'a').setCellStyle(cellStyle.getSimpleCellStyle(true, true, 2, 2, 0, 2));
        sheet.setSameCellStyle(2, 2, 'b', 'e', cellStyle.getSimpleCellStyle(true, true, 2, 2, 1, 1));
        sheet.setSameCellStyle(3, 3, 'a', 'f', cellStyle.getSimpleCellStyle(true, false, 2, 2, 0, 0));
        sheet.setSameCellStyle(9, 9, 'a', 'f', cellStyle.getSimpleCellStyle(true, false, 2, 2, 0, 0));
        sheet.setSameCellStyle(13, 13, 'a', 'f', cellStyle.getSimpleCellStyle(false, false, 2, 0, 0, 0));
        sheet.setSameCellStyle(2, 12, 'g', 'g', cellStyle.getSimpleCellStyle(false, false, 0, 0, 2, 0));
    }

    /**
     * set a wavsep total compare sheet width
     *
     * @param sheet total compare sheet
     */
    private void setTotalCompareWidth(TcSheet sheet) {
        // set a total compare sheet default width column 'A'
        sheet.setColumnWidth('a', POIConstant.WAVSEP_TEST_CASE_CELL_WIDTH);

        // set total compare sheet default width column 'B' ~ 'I'
        for (char i = 'b'; i <= 'i'; i++) {
            sheet.setColumnWidth(i, POIConstant.COMPARE_DEFAULT_COLUMN_WIDTH);
        }
    }

    /**
     * set a wavsep total compare sheet values
     *
     * @param sheet total compare sheet
     */
    private void setTotalCompareValues(TcSheet sheet) {
        sheet.getCell(1, 'a').setCellValue("Compare WAVSEP");
        sheet.getCell(1, 'b').setCellValue(this.date);
        sheet.getCell(2, 'a').setCellValue("Category");
        sheet.getCell(2, 'b').setCellValue("Both True");
        sheet.getCell(2, 'c').setCellValue("Both False");
        sheet.getCell(2, 'd').setCellValue("Only Before True");
        sheet.getCell(2, 'e').setCellValue("Only After True");
        sheet.getCell(2, 'f').setCellValue("Total");
        sheet.getCell(3, 'a').setCellValue("Vulnerabilities + False Positives");

        int row = 4;
        for (Constant.WavsepSheets wavsepSheet : Constant.WavsepSheets.values()) {
            if (!wavsepSheet.toString().equals(Constant.WavsepSheets.total.toString())) {
                sheet.getCell(row, 'a').setCellValue(wavsepSheet.getSheetName());
                sheet.getCell(row, 'b').setCellFormula("\'" + wavsepSheet.getSheetName() + "\'" + "!C5+" + "\'" + wavsepSheet.getSheetName() + "\'" + "!G5");
                sheet.getCell(row, 'c').setCellFormula("\'" + wavsepSheet.getSheetName() + "\'" + "!D5+" + "\'" + wavsepSheet.getSheetName() + "\'" + "!H5");
                sheet.getCell(row, 'd').setCellFormula("\'" + wavsepSheet.getSheetName() + "\'" + "!E5+" + "\'" + wavsepSheet.getSheetName() + "\'" + "!I5");
                sheet.getCell(row, 'e').setCellFormula("\'" + wavsepSheet.getSheetName() + "\'" + "!F5+" + "\'" + wavsepSheet.getSheetName() + "\'" + "!J5");
                sheet.getCell(row, 'f').setCellFormula("SUM(B" + row + ":E" + row + ")");

                row++;
            }
        }
        sheet.getCell(row++, 'a').setCellValue("Experimental Test Cases (inspired / imported from ZAP-WAVE");
        sheet.getCell(row, 'a').setCellValue("additional RXSS(anticsrf tokens, secret input vectors, tag signatures etc.)");
        sheet.getCell(row, 'b').setCellFormula("\'" + Constant.WavsepSheets.rxss.getSheetName() + "\'" + "!K5");
        sheet.getCell(row, 'c').setCellFormula("\'" + Constant.WavsepSheets.rxss.getSheetName() + "\'" + "!L5");
        sheet.getCell(row, 'd').setCellFormula("\'" + Constant.WavsepSheets.rxss.getSheetName() + "\'" + "!M5");
        sheet.getCell(row, 'e').setCellFormula("\'" + Constant.WavsepSheets.rxss.getSheetName() + "\'" + "!N5");
        sheet.getCell(row, 'f').setCellFormula("SUM(B" + row + ":E" + row++ + ")");
        sheet.getCell(row, 'a').setCellValue("addtional SQLi(INSERT)");
        sheet.getCell(row, 'b').setCellFormula("\'" + Constant.WavsepSheets.sqli.getSheetName() + "\'" + "!K5");
        sheet.getCell(row, 'c').setCellFormula("\'" + Constant.WavsepSheets.sqli.getSheetName() + "\'" + "!L5");
        sheet.getCell(row, 'd').setCellFormula("\'" + Constant.WavsepSheets.sqli.getSheetName() + "\'" + "!M5");
        sheet.getCell(row, 'e').setCellFormula("\'" + Constant.WavsepSheets.sqli.getSheetName() + "\'" + "!N5");
        sheet.getCell(row, 'f').setCellFormula("SUM(B" + row + ":E" + row++ + ")");
        sheet.getCell(row, 'a').setCellValue("Totals");
        sheet.getCell(row, 'b').setCellFormula("SUM(B" + 4 + ":B" + (row - 1) + ")");
        sheet.getCell(row, 'c').setCellFormula("SUM(C" + 4 + ":C" + (row - 1) + ")");
        sheet.getCell(row, 'd').setCellFormula("SUM(D" + 4 + ":D" + (row - 1) + ")");
        sheet.getCell(row, 'e').setCellFormula("SUM(E" + 4 + ":E" + (row - 1) + ")");
        sheet.getCell(row, 'f').setCellFormula("SUM(F" + 4 + ":F" + (row - 1) + ")");
    }

    /**
     * set a wavsep vulnerability compare sheet
     *
     * @param sheet vulnerability compare sheet
     */
    private void setVulnerabilityCompareSheet(TcSheet sheet) {
        setVulnerabilityStyles(sheet);
        setVulnerabilityWidth(sheet);
        setVulnerabilityValues(sheet);
    }

    /**
     * set a wavsep vulnerability compare sheet styles
     *
     * @param sheet vulnerability compare sheet
     */
    private void setVulnerabilityStyles(TcSheet sheet) {
        sheet.getSheet().setDefaultColumnStyle(TcUtil.convertColumnAlphabetToIndex('a'), cellStyle.DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);

        for(int i = TcUtil.convertColumnAlphabetToIndex('c'); i <= TcUtil.convertColumnAlphabetToIndex('n'); i++) {
            sheet.getSheet().setDefaultColumnStyle(i, cellStyle.DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        }

        sheet.setSameCellStyle(1, 1, 'a', 'n', cellStyle.getSimpleCellStyle(true, false, 0, 2, 0, 0));
        sheet.setSameCellStyle(2, 4, 'b', 'n', cellStyle.getSimpleCellStyle(true, true, 1, 1, 1, 1));
        sheet.setSameCellStyle(5, 5, 'a', 'n', cellStyle.getSimpleCellStyle(true, true, 1, 2, 1, 1));
    }

    /**
     * set a wavsep vulnerability compare sheet width
     *
     * @param sheet vulnerability compare sheet
     */
    private void setVulnerabilityWidth(TcSheet sheet) {
        sheet.setColumnWidth('a', POIConstant.DEFAULT_COLUMN_WIDTH);
        sheet.setColumnWidth('b', POIConstant.BENCHMARK_TEST_CASE_CELL_WIDTH);
        for (char i = 'c'; i <= 'n'; i++) {
            sheet.setColumnWidth(i, POIConstant.COMPARE_DEFAULT_COLUMN_WIDTH);
        }
    }

    /**
     * set a wavsep vulnerability compare sheet values
     *
     * @param sheet vulnerability compare sheet
     */
    private void setVulnerabilityValues(TcSheet sheet) {
        sheet.getCell(1, 'a').setCellValue(sheet.getSheetName());
        sheet.getCell(5, 'a').setCellValue("합 계");
        sheet.mergeCellRegion(2, 4, 'b', 'b').setCellValue("Test Cases");
        sheet.mergeCellRegion(2, 2, 'c', 'n').setCellValue("Compare");
        sheet.mergeCellRegion(3, 3, 'c', 'f').setCellValue("True Positive");
        sheet.mergeCellRegion(3, 3, 'g', 'j').setCellValue("False Positive");
        sheet.mergeCellRegion(3, 3, 'k', 'n').setCellValue("Experimental");
        sheet.getCell(4, 'c').setCellValue("Both True");
        sheet.getCell(4, 'd').setCellValue("Both False");
        sheet.getCell(4, 'e').setCellValue("Only Before True");
        sheet.getCell(4, 'f').setCellValue("Only After True");
        sheet.getCell(4, 'g').setCellValue("Both True");
        sheet.getCell(4, 'h').setCellValue("Both False");
        sheet.getCell(4, 'i').setCellValue("Only Before True");
        sheet.getCell(4, 'j').setCellValue("Only After True");
        sheet.getCell(4, 'k').setCellValue("Both True");
        sheet.getCell(4, 'l').setCellValue("Both False");
        sheet.getCell(4, 'm').setCellValue("Only Before True");
        sheet.getCell(4, 'n').setCellValue("Only After True");

        sheet.getCell(5, 'b').setCellFormula("COUNTA(B:B)-2");
        sheet.getCell(5, 'c').setCellFormula("COUNTIF(C:C,\"TRUE\")");
        sheet.getCell(5, 'd').setCellFormula("COUNTIF(D:D, \"FALSE\")");
        sheet.getCell(5, 'e').setCellFormula("COUNTIF(E:E, \"TRUE\")");
        sheet.getCell(5, 'f').setCellFormula("COUNTIF(F:F, \"TRUE\")");
        sheet.getCell(5, 'g').setCellFormula("COUNTIF(G:G,\"TRUE\")");
        sheet.getCell(5, 'h').setCellFormula("COUNTIF(H:H, \"FALSE\")");
        sheet.getCell(5, 'i').setCellFormula("COUNTIF(I:I, \"TRUE\")");
        sheet.getCell(5, 'j').setCellFormula("COUNTIF(J:J, \"TRUE\")");
        sheet.getCell(5, 'k').setCellFormula("COUNTIF(K:K,\"TRUE\")");
        sheet.getCell(5, 'l').setCellFormula("COUNTIF(L:L, \"FALSE\")");
        sheet.getCell(5, 'm').setCellFormula("COUNTIF(M:M, \"TRUE\")");
        sheet.getCell(5, 'n').setCellFormula("COUNTIF(N:N, \"TRUE\")");
    }

    /**
     * Write current workbook state in a file
     *
     * @param outputDir output directory path
     * @param fileName  file name
     * @return fale absolute path
     */
    public String writeCompareWorkbook(String outputDir, String fileName) {
        if (compareWorkbook != null) {
            return compareWorkbook.writeWorkbook(outputDir, fileName);
        }
        return null;
    }

    private void compareVulnerabilitySheets(){
        for(Constant.WavsepSheets wavsepSheet : Constant.WavsepSheets.values()){
            if(!wavsepSheet.toString().equals(Constant.WavsepSheets.total.toString())){
                String sheetName = wavsepSheet.getSheetName();
                compareBeforeAndAfterVulnerabilitySheet(compareWorkbook.getTcSheet(sheetName), beforeWorkbook.getTcSheet(sheetName), afterWorkbook.getTcSheet(sheetName));
            }
        }
    }

    /**
     * add before and after total sheet content in compare total sheet
     */
    private void compareTotalSheet(){
        String totalSheetName = Constant.WavsepSheets.total.getSheetName();
        writeBeforeAndAfterTotalInCompareTotal(compareWorkbook.getTcSheet(totalSheetName), beforeWorkbook.getTcSheet(totalSheetName), afterWorkbook.getTcSheet(totalSheetName));
    }

    /**
     * compare before workbook and after workbook
     */
    private void compareBeforeAndAfterVulnerabilitySheet(TcSheet compareSheet, TcSheet beforeSheet, TcSheet afterSheet){
        int tcRow = START_WAVSEP_TC_ROW;
        int compareRow = START_WAVSEP_COMPARE_ROW;

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
                        if (afterSheet.getCell(tcRow, 'h').getStringCellValue().equals("TRUE")) {
                            if (beforeSheet.getCell(tcRow, 'h').getStringCellValue().equals("TRUE")) {
                                /** both true */
                                compareSheet.getCell(compareRow, 'c').setCellValue(true);
                            } else {
                                /** only after true */
                                // beforeSheet.getCell(tcRow, 'i').getStringCellValue().equals("FALSE") same meaning
                                compareSheet.getCell(compareRow, 'f').setCellValue(true);
                            }
                        } else {
                            // afterSheet.getCell(tcRow, 'i').getStringCellValue().equals("FALSE") same meaning
                            if (beforeSheet.getCell(tcRow, 'i').getStringCellValue().equals("FALSE")) {
                                /** both false */
                                compareSheet.getCell(compareRow, 'd').setCellValue(false);
                            } else {
                                /** only before true */
                                // before.getCell(tcRow, 'h').getStringCellValue().equals("TRUE") same meaning
                                compareSheet.getCell(compareRow, 'e').setCellValue(true);
                            }
                        }
                    } else if (!(afterSheet.getCell(tcRow, 'd').getCellType() == Cell.CELL_TYPE_BLANK) && !afterSheet.getCell(tcRow, 'd').getBooleanCellValue()) {
                        /** tc is false positive */
                        if (afterSheet.getCell(tcRow, 'j').getStringCellValue().equals("TRUE")) {
                            if (beforeSheet.getCell(tcRow, 'j').getStringCellValue().equals("TRUE")) {
                                /** both true */
                                compareSheet.getCell(compareRow, 'g').setCellValue(true);
                            } else {
                                /** only after true */
                                // beforeSheet.getCell(tcRow, 'k').getStringCellValue().equals("FALSE") same meaning
                                compareSheet.getCell(compareRow, 'j').setCellValue(true);
                            }
                        } else {
                            // afterSheet.getCell(tcRow, 'k').getStringCellValue().equals("FALSE") same meaning
                            if (beforeSheet.getCell(tcRow, 'k').getStringCellValue().equals("FALSE")) {
                                /** both false */
                                compareSheet.getCell(compareRow, 'h').setCellValue(false);
                            } else {
                                /** only before true */
                                // before.getCell(tcRow, 'j').getStringCellValue().equals("TRUE") same meaning
                                compareSheet.getCell(compareRow, 'i').setCellValue(true);
                            }
                        }
                    } else if (!(afterSheet.getCell(tcRow, 'e').getCellType() == Cell.CELL_TYPE_BLANK) && afterSheet.getCell(tcRow, 'e').getStringCellValue().equals("EX")) {
                        /** tc is experimental */
                        if (afterSheet.getCell(tcRow, 'l').getStringCellValue().equals("TRUE")) {
                            if (beforeSheet.getCell(tcRow, 'l').getStringCellValue().equals("TRUE")) {
                                /** both true */
                                compareSheet.getCell(compareRow, 'k').setCellValue(true);
                            } else {
                                /** only after true */
                                // beforeSheet.getCell(tcRow, 'm').getStringCellValue().equals("FALSE") same meaning
                                compareSheet.getCell(compareRow, 'n').setCellValue(true);
                            }
                        } else {
                            // afterSheet.getCell(tcRow, 'm').getStringCellValue().equals("FALSE") same meaning
                            if (beforeSheet.getCell(tcRow, 'm').getStringCellValue().equals("FALSE")) {
                                /** both false */
                                compareSheet.getCell(compareRow, 'l').setCellValue(false);
                            } else {
                                /** only before true */
                                // before.getCell(tcRow, 'l').getStringCellValue().equals("TRUE") same meaning
                                compareSheet.getCell(compareRow, 'm').setCellValue(true);
                            }
                        }
                    }
                } else {
                    /** after test case and before case is not same */
                    compareCell.setCellStyle(cellStyle.DEFAULT_THICK_RIGHT_LEFT_BG_GRAY);
                }
                // move next row
                tcRow++;
                compareRow++;
            } catch (Exception e){
                e.printStackTrace();
            }
        }
    }

    /**
     * write before and after total sheet contents in copmare total sheet
     * @param compareTotal compare total sheet
     * @param beforeTotal before total sheet
     * @param afterTotal after total sheet
     */
    private void writeBeforeAndAfterTotalInCompareTotal(TcSheet compareTotal, TcSheet beforeTotal, TcSheet afterTotal){
        int compareRow = 16;

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

    private void setBeforeAndAfterTotalStyle(TcSheet target, int startRow){
        int tempRow = startRow;

        CellStyle percentCellStyle = compareWorkbook.getWorkbook().createCellStyle();
        percentCellStyle.setDataFormat(compareWorkbook.getWorkbook().createDataFormat().getFormat("0%"));
        percentCellStyle.setBorderTop(CellStyle.BORDER_THIN);
        percentCellStyle.setBorderBottom(CellStyle.BORDER_THIN);
        percentCellStyle.setBorderLeft(CellStyle.BORDER_THIN);
        percentCellStyle.setBorderRight(CellStyle.BORDER_THIN);

        target.setSameCellStyle(tempRow - 1, tempRow - 1, 'a', 'i', cellStyle.getSimpleCellStyle(true, false, 0, 2, 0, 0));
        target.setSameCellStyle(tempRow, tempRow, 'a', 'i', cellStyle.getSimpleCellStyle(true, true, 1, 1, 1, 1));
        target.setSameCellStyle(++tempRow, tempRow, 'a', 'i', cellStyle.getSimpleCellStyle(true, true, 2, 2, 0 , 0));
        target.setSameCellStyle(++tempRow, tempRow + 4, 'a', 'f', cellStyle.getSimpleCellStyle(false, false, 1, 1, 1, 1));
        target.setSameCellStyle(tempRow, tempRow + 4, 'g', 'i', percentCellStyle);
        tempRow = tempRow + 4;
        target.setSameCellStyle(++tempRow, tempRow, 'a', 'i', cellStyle.getSimpleCellStyle(true, true, 2, 2, 0, 0));
        target.setSameCellStyle(++tempRow, tempRow + 3, 'a', 'f', cellStyle.getSimpleCellStyle(false, false, 1, 1, 1, 1));
        target.setSameCellStyle(tempRow, tempRow + 3, 'g', 'i', percentCellStyle);
        tempRow = tempRow + 3;
        target.setSameCellStyle(startRow, tempRow, 'j', 'j', cellStyle.getSimpleCellStyle(false, false, 0, 0, 2, 0));
        target.setSameCellStyle(++tempRow, tempRow, 'a', 'i', cellStyle.getSimpleCellStyle(false, false, 2, 0, 0, 0));
    }

    public static void main(String[] args){
        try {
            WavsepCompareTemplate compare = new WavsepCompareTemplate("C:\\scalaProjects\\testPOI\\Zap_Wavsep_Results_160602140051.xlsx",
                    "C:\\scalaProjects\\testPOI\\Zap_Wavsep_Results_160608161934.xlsx");

            compare.writeCompareWorkbook("./", "test.xlsx");
        } catch (IOException e){
            e.printStackTrace();
        }
    }
}
