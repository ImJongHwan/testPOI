package poi.TestCasePOI;

import org.apache.poi.ss.usermodel.Cell;
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

    public Cell getCell(int rowNum, char columnAlphabet){
        char upperColumnAlphabet = Character.toUpperCase(columnAlphabet);

        TcRow tcRow = getTcRow(rowNum);
        Cell cell = tcRow.getCell(upperColumnAlphabet);

        return cell;
    }

    public XSSFSheet getSheet() {
        return sheet;
    }
}
