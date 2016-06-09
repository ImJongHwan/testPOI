package org.poi.TestCasePOI;

import org.poi.POIConstant;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;

import java.util.HashMap;
import java.util.Map;

/**
 * Created by Hwan on 2016-05-10.
 */
public class TcRow {

    private XSSFRow row = null;
    private static final int ALPHABET_SIZE = 26;

    TcRow(XSSFRow row) {
        this.row = row;
    }

    /**
     * Create and Get cell in row.
     *
     * @param columnAlphabet columnAlphabet in real excel.
     * @return cell
     */
    Cell getCell(String columnAlphabet) {
        return getCell(convertColumnAlphabetToIndex(columnAlphabet));
    }

    /**
     * Create and get cell in row
     *
     * @param columnAlphabet A single columnAlphabet in real excel.
     * @return cell
     */
    Cell getCell(char columnAlphabet) {
        String charToString = String.valueOf(columnAlphabet);
        return getCell(charToString);
    }

    /**
     * Get cell in row by columnIndex
     *
     * @param columnIndex XSSFCell ColumnIndex
     * @return cell
     */
    Cell getCell(int columnIndex) {
        if(row.getCell(columnIndex) != null){
            return row.getCell(columnIndex);
        } else {
            return row.createCell(columnIndex);
        }
    }

    /**
     * Get cell apart column offset
     *
     * @param columnAlphabet base column alphabet
     * @param offset         distance
     * @return cell
     */
    Cell getOffsetCell(String columnAlphabet, int offset) {
        return getCell(convertColumnAlphabetToIndex(columnAlphabet) + offset);
    }

    /**
     * Get column alphabet apart column offset
     *
     * @param columnAlphabet base column alphabet
     * @param offset         distance
     * @return column alphabet
     */
    public static String getOffsetColumnAlphabet(String columnAlphabet, int offset) {
        return convertColumnIndexToAlphabet(convertColumnAlphabetToIndex(columnAlphabet) + offset);
    }

    /**
     * Convert to column index to column alphabet
     *
     * @param columnIndex XSSFCell column index
     * @return columnAlphabet
     */
    private static String convertColumnIndexToAlphabet(int columnIndex) {
        if(columnIndex == 0){
            return "A";
        }

        String columnAlphabet = "";

        int tempIndex;
        int resIndex = columnIndex;

        while (resIndex > 0) {
            tempIndex = (resIndex % ALPHABET_SIZE) + 'A';
            columnAlphabet = String.valueOf(Character.toChars(tempIndex)) + columnAlphabet;
            resIndex = resIndex / ALPHABET_SIZE;
        }

        return columnAlphabet;
    }

    /**
     * convert ColumnAlphabet in real excel to Cell Index
     *
     * @param columnAlphabet columnAlphabet in real excel
     * @return cell index
     */
    private static int convertColumnAlphabetToIndex(String columnAlphabet) {
//        columnAlphabet = columnAlphabet.toUpperCase();

        int cellIndex = 0;

        for (int i = 0; i < columnAlphabet.length(); i++) {
            cellIndex += (columnAlphabet.charAt(i) - 'A') * Math.pow(ALPHABET_SIZE, columnAlphabet.length() - i - 1);
        }

        return cellIndex;
    }

    /**
     * Check whether s is contained integer.
     *
     * @param s string
     * @return is contained integer
     */
    private static boolean isInteger(String s) {
        try {
            Integer.parseInt(s);
        } catch (NumberFormatException e) {
            return false;
        } catch (NullPointerException e) {
            return false;
        }
        return true;
    }


    XSSFRow getRow() {
        return row;
    }
}
