package org.poi.TestCasePOI;

import org.poi.Util.TcUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.poi.POIConstant;

import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;

/**
 * Created by Hwan on 2016-05-10.
 */
public class TcSheet {
    private XSSFWorkbook parentWorkbook = null;
    private String sheetName = null;

    private XSSFSheet sheet = null;
    private int tpTcEndRowNum = 0;
    private int fpTcEndRowNum = 0;
    private int exTcEndRowNum = 0;

    private Map<Integer, TcRow> rowMap = new HashMap<>();

    public TcSheet(XSSFWorkbook parentWorkbook, String sheetName) {
        this.parentWorkbook = parentWorkbook;
        this.sheetName = sheetName;
        this.sheet = this.parentWorkbook.createSheet(sheetName);
    }

    public TcSheet(XSSFWorkbook parentWorkbook, XSSFSheet sheet) {
        this.parentWorkbook = parentWorkbook;
        this.sheetName = sheet.getSheetName();
        this.sheet = sheet;
    }

    /**
     * Get row
     * ! Caution : parameter rowNum is not XSSFRow index num but real excel row line num.
     *
     * @param rowNum row line number (real excel row line number)
     * @return row
     */
    public XSSFRow getRow(int rowNum) {
        return getTcRow(rowNum).getParentRow();
    }

    /**
     * Get a TcRow instant
     * ! Caution : parameter rowNum is not XSSFRow index num but real excel row line num.
     *
     * @param rowNum row line number (real excel row line number)
     * @return A TcRow instant
     */
    public TcRow getTcRow(int rowNum) {
        if (rowNum < 1) {
            System.err.println("Getting Row Error - A row number is negative. Please check a row number.");
        }

        if (rowMap.containsKey(rowNum)) {
            return rowMap.get(rowNum);
        }

        // excel row line number start 1 but XSSFRow index start 0
        XSSFRow row = sheet.createRow(rowNum - 1);
        TcRow tcRow = new TcRow(row);

        rowMap.put(rowNum, tcRow);

        return tcRow;
    }

    /**
     * Create and get cell.
     *
     * @param rowNum         row line number (real excel row line number)
     * @param columnAlphabet column alphabet (real excel column alphabet)
     * @return cell
     */
    public Cell getCell(int rowNum, String columnAlphabet) {
        String upperColumnAlphabet = columnAlphabet.toUpperCase();

        TcRow tcRow = getTcRow(rowNum);
        Cell cell = tcRow.getCell(upperColumnAlphabet);

        return cell;
    }

    /**
     * Create and get cell
     *
     * @param rowNum         row line number (real excel row line number)
     * @param columnAlphabet A single column alphabet (real excel column alphabet)
     * @return cell
     */
    public Cell getCell(int rowNum, char columnAlphabet) {
        char upperColumnAlphabet = Character.toUpperCase(columnAlphabet);

        TcRow tcRow = getTcRow(rowNum);
        Cell cell = tcRow.getCell(upperColumnAlphabet);

        return cell;
    }

    /**
     * Create and get Cell by XSSFCell row index and column index
     *
     * @param rowIndex    XSSFCell row index
     * @param columnIndex XSSFCell column index
     * @return cell
     */
    public Cell getCell(int rowIndex, int columnIndex) {
        TcRow tcRow = getTcRow(rowIndex + 1);
        Cell cell = tcRow.getCell(columnIndex);

        return cell;
    }

    /**
     * merge cells region and get a first cell
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn column containing first cell
     * @param lastColumn  column containing last cell
     * @return first cell in merged cell region
     */
    public Cell mergeCellRegion(int firstRow, int lastRow, String firstColumn, String lastColumn) {

        // Error check -  Must be lastRow > firstRow
        if (firstRow > lastRow) {
            System.err.println("Merge row data Error - Last row data must be bigger than first row data.");
            return null;
        }

        // Error check - Must be lastColumn > firstColumn
        if (firstColumn.compareTo(lastColumn) > 0) {
            System.err.println("Merge column data Error - Last Column data must be bigger than first column data.");
            return null;
        }

        if ((firstRow == lastRow) && (firstColumn.compareTo(lastColumn) == 0)) {
            System.out.println("First row data = Last row data, and first column data = last column data - It is a only one cell.");
            return getCell(firstRow, firstColumn);
        }

        Cell firstCell = getCell(firstRow, firstColumn);
        Cell lastCell = getCell(lastRow, lastColumn);

        Cell firstCellExistedMerged;

        // First cell is already contained merged region.
        if ((firstCellExistedMerged = getFirstCell(firstCell)) != null) {
            System.out.println("This first cell is already contained merged region.");
            return firstCellExistedMerged;
        }

        // Last cell is already contained merged region.
        if ((firstCellExistedMerged = getFirstCell(lastCell)) != null) {
            System.out.println("This last cell is already contained merged region.");
            return firstCellExistedMerged;
        }

        CellRangeAddress mergedRegion = new CellRangeAddress(firstCell.getRowIndex(), lastCell.getRowIndex(),
                firstCell.getColumnIndex(), lastCell.getColumnIndex());

        sheet.addMergedRegion(mergedRegion);

        return firstCell;
    }

    /**
     * merge cells and get a first cell
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn a single column alphabet containing a first cell
     * @param lastColumn  a single column alphabet containing a last cell
     * @return first cell in merged cell region
     */
    public Cell mergeCellRegion(int firstRow, int lastRow, char firstColumn, char lastColumn) {
        String firstColumnString = String.valueOf(firstColumn);
        String lastColumnString = String.valueOf(lastColumn);

        return mergeCellRegion(firstRow, lastRow, firstColumnString, lastColumnString);
    }

    /**
     * get cell range address to use real row int, column characters
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn a single column alphabet containing a first cell
     * @param lastColumn  a single column alphabet containing a last cell
     * @return CellRangeAddress
     */
    public CellRangeAddress getCellRange(int firstRow, int lastRow, String firstColumn, String lastColumn) {
        // Error check -  Must be lastRow > firstRow
        if (firstRow > lastRow) {
            System.err.println("Merge row data Error - Last row data must be bigger than first row data.");
            return null;
        }

        // Error check - Must be lastColumn > firstColumn
        if (firstColumn.compareTo(lastColumn) > 0) {
            System.err.println("Merge column data Error - Last Column data must be bigger than first column data.");
            return null;
        }

        if ((firstRow == lastRow) && (firstColumn.compareTo(lastColumn) == 0)) {
            System.out.println("First row data = Last row data, and first column data = last column data - It is a only one cell.");
            return null;
        }

        Cell firstCell = getCell(firstRow, firstColumn);
        Cell lastCell = getCell(lastRow, lastColumn);

        CellRangeAddress cellRange = new CellRangeAddress(firstCell.getRowIndex(), lastCell.getRowIndex(),
                firstCell.getColumnIndex(), lastCell.getColumnIndex());

        return cellRange;
    }

    /**
     * get cell range address to use real row int, column characters
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn a single column alphabet containing a first cell
     * @param lastColumn  a single column alphabet containing a last cell
     * @return CellRangeAddress
     */
    public CellRangeAddress getCellRange(int firstRow, int lastRow, char firstColumn, char lastColumn) {
        String firstColumnString = String.valueOf(firstColumn);
        String lastColumnString = String.valueOf(lastColumn);

        return getCellRange(firstRow, lastRow, firstColumnString, lastColumnString);
    }

    /**
     * Get a first cell in merged region containing this cell.
     *
     * @param cell a cell contained merged region
     * @return first cell
     */
    public Cell getFirstCell(Cell cell) {
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress region = sheet.getMergedRegion(i);
            if (region.getFirstRow() <= cell.getRowIndex() && region.getLastRow() >= cell.getRowIndex()) {
                if (region.getFirstColumn() <= cell.getColumnIndex() && region.getLastColumn() >= cell.getColumnIndex()) {
                    Cell cellInMerged = getCell(region.getFirstRow(), region.getFirstColumn());
                    return cellInMerged;
                }
            }
        }
        return null;
    }

    /**
     * Multi cells set to be same value.
     * ! caution : value must be String or Integer.
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn a single column alphabet containing a first cell
     * @param lastColumn  a single column alphabet containing a last cell
     * @param value       value (It must be String or Integer)
     */
    public void setSameValueMultiCells(int firstRow, int lastRow, String firstColumn, String lastColumn, Object value) {

        String upperFirst = firstColumn.toUpperCase();
        String upperLast = lastColumn.toUpperCase();

        for (int i = firstRow; i <= lastRow; i++) {
            TcRow tcRow = getTcRow(i);

            String tempColumn = upperFirst;
            while (!(tempColumn.compareTo(upperLast) > 0)) {

                Cell tempCell = getCell(i, tempColumn);

                if (value instanceof String) {
                    tempCell.setCellValue(value.toString());
                } else if (value instanceof Integer) {
                    tempCell.setCellValue((int) value);
                } else if (value instanceof Boolean) {
                    tempCell.setCellValue((boolean) value);
                } else if (value instanceof Calendar) {
                    tempCell.setCellValue((Calendar) value);
                } else if (value instanceof Date) {
                    tempCell.setCellValue((Date) value);
                } else if (value instanceof RichTextString) {
                    tempCell.setCellValue((RichTextString) value);
                } else {
                    System.err.println("Set value Error - Setting multi cells to be same value is failed.");
                    return;
                }

                tempColumn = tcRow.getOffsetColumnAlphabet(tempColumn, 1);
            }
        }
    }

    public void setSameValueMultiCells(int firstRow, int lastRow, char firstColumn, char lastColumn, Object value) {
        String firstColumnString = String.valueOf(firstColumn);
        String lastColumnString = String.valueOf(lastColumn);

        setSameValueMultiCells(firstRow, lastRow, firstColumnString, lastColumnString, value);
    }

    public XSSFSheet getSheet() {
        return sheet;
    }

    public String getSheetName() {
        return sheetName;
    }

    /**
     * set column width to use real column alphabet
     * !caution : setColumnWidth API has a error - you must multiply a POIConstant.COLUMN_SIZE_CONSTANT(256)
     *
     * @param columnIndex column alphabet
     * @param width       column width
     */
    public void setColumnWidth(String columnIndex, int width) {
        this.sheet.setColumnWidth(TcUtil.convertColumnAlphabetToIndex(columnIndex), POIConstant.COLUMN_SIZE_CONSTANT * width);
    }

    /**
     * set column width to use real column alphabet
     * !caution : setColumnWidth API has a error - you must multiply a POIConstant.COLUMN_SIZE_CONSTANT(256)
     *
     * @param columnIndex single column alphabet
     * @param width       column width
     */
    public void setColumnWidth(char columnIndex, int width) {
        this.sheet.setColumnWidth(TcUtil.convertColumnAlphabetToIndex(columnIndex), POIConstant.COLUMN_SIZE_CONSTANT * width);
    }

    /**
     * set style to multi cells
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn a single column alphabet containing a first cell
     * @param lastColumn  a single column alphabet containing a last cell
     * @param style       cell style
     */
    public void setSameCellStyle(int firstRow, int lastRow, String firstColumn, String lastColumn, CellStyle style) {

        String upperFirst = firstColumn.toUpperCase();
        String upperLast = lastColumn.toUpperCase();

        if (style == null) {
            System.out.println("Set style error - style is a null character : TcSheet");
            return;
        }

        for (int i = firstRow; i <= lastRow; i++) {
            TcRow tcRow = getTcRow(i);

            String tempColumn = upperFirst;
            while (!(tempColumn.compareTo(upperLast) > 0)) {

                Cell tempCell = getCell(i, tempColumn);

                tempCell.setCellStyle(style);

                tempColumn = tcRow.getOffsetColumnAlphabet(tempColumn, 1);
            }
        }
    }

    /**
     * set style to multi cells
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn a single column alphabet containing a first cell
     * @param lastColumn  a single column alphabet containing a last cell
     * @param style       cell style
     */
    public void setSameCellStyle(int firstRow, int lastRow, char firstColumn, char lastColumn, CellStyle style) {
        String firstColumnString = String.valueOf(firstColumn);
        String lastColumnString = String.valueOf(lastColumn);

        setSameCellStyle(firstRow, lastRow, firstColumnString, lastColumnString, style);
    }

    /**
     * set same formula to multi cells
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn column alphabet containing a first cell
     * @param lastColumn  column alphabet containing a last cell
     * @param formula     formula
     */
    public void setSameFormula(int firstRow, int lastRow, String firstColumn, String lastColumn, String formula) {
        String upperFirst = firstColumn.toUpperCase();
        String upperLast = lastColumn.toUpperCase();

        if (formula == null) {
            System.out.println("Set style error - style is a null character : TcSheet");
            return;
        }

        for (int i = firstRow; i <= lastRow; i++) {
            TcRow tcRow = getTcRow(i);

            String tempColumn = upperFirst;
            while (!(tempColumn.compareTo(upperLast) > 0)) {

                Cell tempCell = getCell(i, tempColumn);

                tempCell.setCellFormula(formula);

                tempColumn = tcRow.getOffsetColumnAlphabet(tempColumn, 1);
            }
        }
    }

    /**
     * set same formula to multi cells
     *
     * @param firstRow    row containing first cell
     * @param lastRow     row containing last cell
     * @param firstColumn column alphabet containing a first cell
     * @param lastColumn  column alphabet containing a last cell
     * @param formula     formula
     */
    public void setSameFormula(int firstRow, int lastRow, char firstColumn, char lastColumn, String formula) {
        String firstColumnString = String.valueOf(firstColumn);
        String lastColumnString = String.valueOf(lastColumn);

        setSameFormula(firstRow, lastRow, firstColumnString, lastColumnString, formula);
    }

    /**
     * get true positive test cases end row number
     *
     * @return true positive test cases end row number
     */
    public int getTpTcEndRowNum() {
        return tpTcEndRowNum;
    }

    /**
     * set true positive test cases end row number
     *
     * @param tpTcEndRowNum true positive test cases end row number
     */
    public void setTpTcEndRowNum(int tpTcEndRowNum) {
        this.tpTcEndRowNum = tpTcEndRowNum;
    }

    /**
     * get false positive test cases end row number
     *
     * @return false positive test cases end row number
     */
    public int getFpTcEndRowNum() {
        return fpTcEndRowNum;
    }

    /**
     * set false positive test casea end row number
     *
     * @param fpTcEndRowNum false positive test cases end row number
     */
    public void setFpTcEndRowNum(int fpTcEndRowNum) {
        this.fpTcEndRowNum = fpTcEndRowNum;
    }

    /**
     * get experimental test cases end row number
     *
     * @return experimental test cases end row number
     */
    public int getExTcEndRowNum() {
        return exTcEndRowNum;
    }

    /**
     * set experimental test cases end row number
     *
     * @param exTcEndRowNum experimental test cases end row number
     */
    public void setExTcEndRowNum(int exTcEndRowNum) {
        this.exTcEndRowNum = exTcEndRowNum;
    }
}