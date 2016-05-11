package poi.TestCasePOI;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by Hwan on 2016-05-10.
 */
public class TcRow {

    private XSSFRow parentRow = null;
    private Map<String, Cell> cellMap = new HashMap<>();
    private static final int ALPHABET_SIZE = 26;

    TcRow(XSSFRow parentRow){
        this.parentRow = parentRow;
    }

    /**
     * Create and Get cell in parentRow.
     *
     * @param columnAlphabet columnAlphabet in real excel.
     * @return cell
     */
    Cell getCell(String columnAlphabet){
        if(columnAlphabet.compareTo("XFD") > 0){
            System.err.println("Getting cell Error - The columnAlphabet is too big. Please input a columnAlphabet less than \"XFD\".");
            return null;
        } else if (isInteger(columnAlphabet)){
            System.err.println("ColumnAlphabet Error - The ColumnAlphabet is contained number. Please check columnAlphabet.");
            return null;
        }

        if(cellMap.containsKey(columnAlphabet)){
            return cellMap.get(columnAlphabet);
        }

        Cell cell = parentRow.createCell(convertColumnAlphabetToIndex(columnAlphabet));

        cellMap.put(columnAlphabet, cell);

        return cell;
    }

    /**
     * Create and get cell in parentRow
     *
     * @param columnAlphabet A single columnAlphabet in real excel.
     * @return cell
     */
    Cell getCell(char columnAlphabet){
        if(!Character.isAlphabetic(columnAlphabet)){
            System.err.println("Getting cell Error - The Single columnAlphabet is number. Please input a alphabet.");
            return null;
        }

        String charToString = String.valueOf(columnAlphabet);

        if(cellMap.containsKey(charToString)){
            return cellMap.get(charToString);
        }

        Cell cell = parentRow.createCell(convertColumnAlphabetToIndex(charToString));

        cellMap.put(charToString, cell);

        return cell;
    }

    /**
     * convert ColumnAlphabet in real excel to Cell Index
     *
     * @param columnAlphabet columnAlphabet in real excel
     * @return cell index
     */
    private int convertColumnAlphabetToIndex(String columnAlphabet){
//        columnAlphabet = columnAlphabet.toUpperCase();

        int cellIndex = 0;

        for(int i = 0; i < columnAlphabet.length(); i++){
            cellIndex += (columnAlphabet.charAt(i)  -'A') *  Math.pow( ALPHABET_SIZE ,columnAlphabet.length() - i - 1);
        }

        return cellIndex;
    }

    /**
     * Check whether s is contained integer.
     *
     * @param s string
     * @return is contained integer
     */
    private static boolean isInteger(String s){
        try {
            Integer.parseInt(s);
        } catch (NumberFormatException e) {
            return false;
        } catch (NullPointerException e){
            return false;
        }
        return true;
    }


    XSSFRow getParentRow() {
        return parentRow;
    }
}
