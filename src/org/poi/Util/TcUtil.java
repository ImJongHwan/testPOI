package org.poi.Util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.RichTextString;
import org.POIConstant;
import org.poi.TestCasePOI.TcSheet;

import java.util.*;

/**
 * Created by Hwan on 2016-05-26.
 */
public class TcUtil {

    /**
     * convert ColumnAlphabet in real excel to Cell Index
     *
     * @param columnAlphabet columnAlphabet in real excel
     * @return cell index
     */
    public static int convertColumnAlphabetToIndex(String columnAlphabet) {
        columnAlphabet = columnAlphabet.toUpperCase();

        int cellIndex = 0;

        for (int i = 0; i < columnAlphabet.length(); i++) {
            cellIndex += (columnAlphabet.charAt(i) - 'A') * Math.pow(POIConstant.ALPHABET_SIZE, columnAlphabet.length() - i - 1);
        }

        return cellIndex;
    }

    /**
     * convert ColumnAlphabet in real excel to Cell Index
     *
     * @param columnAlphabet columnAlphabet in real excel
     * @return cell index
     */
    public static int convertColumnAlphabetToIndex(char columnAlphabet) {
        String columnString = String.valueOf(columnAlphabet);

        return convertColumnAlphabetToIndex(columnString);
    }

    /**
     * write down data list
     *
     * @param list list to write
     * @param tcSheet tcSheet to write in
     * @param startRow start cell Row
     * @param startColumn start cell Column
     * @param <T> input data type : String, Calendar, Date, Double, Boolean, RichTextString
     * @Return end Row + 1
     *
     */
    public static <T> int writeDownListInSheet(List<T> list, TcSheet tcSheet, int startRow, String startColumn){
        if(list == null || tcSheet == null || list.isEmpty()){
            return -1;
        }

        int rowNum = startRow;

        for(T content : list){
            Cell cell = tcSheet.getCell(rowNum, startColumn);

            if(content instanceof String) {
                cell.setCellValue((String) content);
            } else if (content instanceof Calendar){
                cell.setCellValue((Calendar) content);
            } else if (content instanceof Date){
                cell.setCellValue((Date) content);
            } else if (content instanceof Double){
                cell.setCellValue((Double) content);
            } else if (content instanceof Boolean){
                cell.setCellValue((Boolean) content);
            } else if (content instanceof RichTextString){
                cell.setCellValue((RichTextString) content);
            } else {
                System.out.println("List contents are invalid data type - TcUtil" );
                return -1;
            }

            rowNum++;
        }

        return rowNum;
    }

    /**
     * write down list
     *
     * @param list list to write
     * @param tcSheet tcSheet to write in
     * @param startRow start cell Row
     * @param startColumn start cell Column
     * @param <T> input data type : String, Calendar, Date, double, boolean, RichTextString
     */
    public static <T> int writeDownListInSheet(List<T> list, TcSheet tcSheet, int startRow, char startColumn){
        String startColumnString = String.valueOf(startColumn);

        return writeDownListInSheet(list, tcSheet, startRow, startColumnString);
    }
}
