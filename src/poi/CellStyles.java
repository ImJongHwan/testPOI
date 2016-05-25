package poi;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * Created by Hwan on 2016-05-24.
 */
public class CellStyles {
    /**
     * get cell style that set bold font and align center.
     *
     * @param workbook workbook to set style
     * @return CellStyle
     */
    public static CellStyle getBoldCenterStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(getBoldFont(workbook));

        cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
        cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);

        return cellStyle;
    }

    /**
     * get cell style that set only bold.
     *
     * @param workbook workbook to set style
     * @return CellStyle
     */
    public static CellStyle getBoldStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(getBoldFont(workbook));

        return cellStyle;
    }

    public static CellStyle setDefaultStyle(Workbook workbook){
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(setDefaultFont(workbook));

        return cellStyle;
    }

    /**
     * get font style that set only bold
     *
     * @param workbook workbook to set font style
     * @return Font
     */
    private static Font getBoldFont(Workbook workbook) {
        Font font = setDefaultFont(workbook);
        font.setBold(true);

        return font;
    }

    /**
     * get default font style
     * Font name : "맑은 고딕"
     * Font height : 11
     *
     * @param workbook workbook to set font style
     * @return Font
     */
    private static Font setDefaultFont(Workbook workbook){
        Font defaultFont = workbook.createFont();

        defaultFont.setFontName(Constant.DEFAULT_FONT_NAME);

        short defaultFontHeight = Constant.DEFAULT_FONT_HEIGHT * Constant.FONT_SIZE_CONSTANT;
        defaultFont.setFontHeight(defaultFontHeight);

        return defaultFont;
    }

    /**
     * append Style Top Border
     *
     * @param cellStyle cellStyle
     * @param borderStyle borderStyle, you can get a CellStyle class.
     */
    public static void appendTopBorder(CellStyle cellStyle, short borderStyle){
        if(cellStyle != null) {
            cellStyle.setBorderTop(borderStyle);
        }
    }

    /**
     * append Style Bottom Border
     *
     * @param cellStyle cellStyle
     * @param borderStyle borderStyle, you can get a CellStyle class.
     */
    public static void appendBottomBorder(CellStyle cellStyle, short borderStyle){
        if(cellStyle != null) {
            cellStyle.setBorderRight(borderStyle);
        }
    }

    /**
     * append Style Right Border
     *
     * @param cellStyle cellStyle
     * @param borderStyle borderStyle, you can get a CellStyle class.
     */
    public static void appendRightBorder(CellStyle cellStyle, short borderStyle){
        if(cellStyle != null) {
            cellStyle.setBorderRight(borderStyle);
        }
    }

    /**
     * append Style Left Border
     *
     * @param cellStyle cellStyle
     * @param borderStyle borderStyle, you can get a CellStyle class.
     */
    public static void appendLeftBorder(CellStyle cellStyle, short borderStyle){
        if(cellStyle != null) {
            cellStyle.setBorderLeft(borderStyle);
        }
    }

}