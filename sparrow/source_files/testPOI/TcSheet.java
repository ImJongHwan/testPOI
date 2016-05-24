package poi.TestCasePOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by Hwan on 2016-05-10.
 */
public class TcSheet {
    private XSSFWorkbook parentWorkbook = null;
    private String sheetName = null;

    private XSSFSheet sheet = null;

    private Map<Integer, TcRow> rowMap = new HashMap<>();

    public TcSheet(XSSFWorkbook parentWorkbook, String sheetName){
        this.parentWorkbook = parentWorkbook;
        this.sheetName = sheetName;
        this.sheet = this.parentWorkbook.createSheet(sheetName);
    }

    public TcSheet(XSSFWorkbook parentWorkbook, XSSFSheet sheet){
        this.parentWorkbook = parentWorkbook;
        this.sheetName = sheet.getSheetName();
        this.sheet = sheet;
    }

    /**
     * Get row
     *! Caution : parameter rowNum is not XSSFRow index num but real excel row line num.
     *
     * @param rowNum row line number (real excel row line number)
     * @return row
     */
    public XSSFRow getRow(int rowNum){
        return getTcRow(rowNum).getParentRow();
    }

    /**
     * Get a TcRow instant
     * ! Caution : parameter rowNum is not XSSFRow index num but real excel row line num.
     *
     * @param rowNum row line number (real excel row line number)
     * @return A TcRow instant
     */
    public TcRow getTcRow(int rowNum){
        if(rowNum < 1){
            System.err.println("Getting Row Error - A row number is negative. Please check a row number.");
        }

        if(rowMap.containsKey(rowNum)){
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
     * @param rowNum row line number (real excel row line number)
     * @param columnAlphabet column alphabet (real excel column alphabet)
     * @return cell
     */
    public Cell getCell(int rowNum, String columnAlphabet){
        String upperColumnAlphabet =columnAlphabet.toUpperCase();

        TcRow tcRow = getTcRow(rowNum);
        Cell cell = tcRow.getCell(upperColumnAlphabet);

        return cell;
    }

    /**
     * Create and get cell
     *
     * @param rowNum row line number (real excel row line number)
     * @param columnAlphabet A single column alphabet (real excel column alphabet)
     * @return cell
     */
    public Cell getCell(int rowNum, char columnAlphabet){
        char upperColumnAlphabet = Character.toUpperCase(columnAlphabet);

        TcRow tcRow = getTcRow(rowNum);
        Cell cell = tcRow.getCell(upperColumnAlphabet);

        return cell;
    }

    /**
     * Create and get Cell by XSSFCell row index and column index
     *
     * @param rowIndex XSSFCell row index
     * @param columnIndex XSSFCell column index
     * @return cell
     */
    public Cell getCell(int rowIndex, int columnIndex){
        TcRow tcRow = getTcRow(rowIndex + 1);
        Cell cell = tcRow.getCell(columnIndex);

        return cell;
    }

    /**
     * merge cells region and get a first cell
     *
     * @param firstRow row containing first cell
     * @param lastRow row containing last cell
     * @param firstColumn column containing first cell
     * @param lastColumn column containing last cell
     * @return first cell in merged cell region
     */
    public Cell mergeCellRegion(int firstRow, int lastRow, String firstColumn, String lastColumn){

        // Error check -  Must be lastRow > firstRow
        if(firstRow > lastRow){
            System.err.println("Merge row data Error - Last row data must be bigger than first row data.");
            return null;
        }

        // Error check - Must be lastColumn > firstColumn
        if(firstColumn.compareTo(lastColumn) > 0){
            System.err.println("Merge column data Error - Last Column data must be bigger than first column data.");
            return null;
        }

        if((firstRow == lastRow) && (firstColumn.compareTo(lastColumn) == 0)){
            System.out.println("First row data = Last row data, and first column data = last column data - It is a only one cell.");
            return getCell(firstRow, firstColumn);
        }

        Cell firstCell = getCell(firstRow, firstColumn);
        Cell lastCell = getCell(lastRow, lastColumn);

        Cell firstCellExistedMerged;

        // First cell is already contained merged region.
        if((firstCellExistedMerged = getFirstCell(firstCell)) != null){
            System.out.println("This first cell is already contained merged region.");
            return firstCellExistedMerged;
        }

        // Last cell is already contained merged region.
        if((firstCellExistedMerged = getFirstCell(lastCell)) != null){
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
     * @param firstRow row containing first cell
     * @param lastRow row containing last cell
     * @param firstColumn a single column alphabet containing a first cell
     * @param lastColumn a single column alphabet containing a last cell
     * @return first cell in merged cell region
     */
    public Cell mergeCellRegion(int firstRow, int lastRow, char firstColumn, char lastColumn){
        String firstColumnString = String.valueOf(firstColumn);
        String lastColumnString = String.valueOf(lastColumn);

        return mergeCellRegion(firstRow, lastRow, firstColumnString, lastColumnString);
    }

    /**
     * Get a first cell in merged region containing this cell.
     *
     * @param cell a cell contained merged region
     * @return first cell
     */
    public Cell getFirstCell(Cell cell){
        for(CellRangeAddress region : sheet.getMergedRegions()){
            if(region.getFirstRow() <= cell.getRowIndex() && region.getLastRow() >= cell.getRowIndex()){
                if(region.getFirstColumn() <= cell.getColumnIndex() && region.getLastColumn() >= cell.getColumnIndex()){
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
     * @param firstRow row containing first cell
     * @param lastRow row containing last cell
     * @param firstColumn a single column alphabet containing a first cell
     * @param lastColumn a single column alphabet containing a last cell
     * @param value value (It must be String or Integer)
     */
    public void setSameValueMultiCells(int firstRow, int lastRow, String firstColumn, String lastColumn, Object value){
        for(int i = firstRow; i <= lastRow; i++){
            TcRow tcRow = getTcRow(i);

            String tempColumn = firstColumn;
            while(!(tempColumn.compareTo(lastColumn)>0)){

                Cell tempCell = getCell(i, tempColumn);

                if(value instanceof String){
                    tempCell.setCellValue(value.toString());
                } else if(value instanceof Integer){
                    tempCell.setCellValue((int) value);
                } else {
                    System.err.println("Set value Error - Setting multi cells to be same value is failed. Value must be String or Integer.");
                    return;
                }

                tempColumn = tcRow.getOffsetColumnAlphabet(tempColumn,1);
            }
        }
    }

    public XSSFSheet getSheet() {
        return sheet;
    }

    public String getSheetName() {
        return sheetName;
    }
}