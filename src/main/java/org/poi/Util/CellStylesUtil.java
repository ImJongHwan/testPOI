package org.poi.Util;

import org.poi.POIConstant;
import org.apache.poi.ss.usermodel.*;

/**
 * Created by Hwan on 2016-05-24.
 */
public class CellStylesUtil {

    public CellStyle BOLD_CENTER_THICK_TOP_BOTTOM;
    public CellStyle BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT;
    public CellStyle BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT;
    public CellStyle BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT;
    public CellStyle BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT;
    public CellStyle BOLD_CENTER_MIDDLE_RIGHT_LEFT;

    public CellStyle BOLD_DEFAULT_THICK_TOP_BOTTOM;

    public CellStyle DEFAULT_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT;
    public CellStyle DEFAULT_CENTER_THICK_BOTTOM_MIDDLE_RIGHT_LEFT;
    public CellStyle DEFAULT_CENTER_MIDDLE_RIGHT_LEFT;

    public CellStyle DEFAULT_DEFAULT_THICK_TOP_MIDDLE_RIGHT_LEFT;
    public CellStyle DEFAULT_DEFAULT_THICK_BOTTOM_MIDDLE_RIGHT_LEFT;
    public CellStyle DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT;

    public CellStyle DEFAULT_THICK_RIGHT_LEFT_BG_GRAY;

    private Workbook workbook = null;

    public CellStylesUtil(Workbook workbook) {
        this.workbook = workbook;
        init(workbook);
    }

    private void init(Workbook workbook) {
        BOLD_CENTER_THICK_TOP_BOTTOM = getSimpleCellStyle(true, true, 2, 2, 0, 0);

        BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT.cloneStyleFrom(BOLD_CENTER_THICK_TOP_BOTTOM);
        BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT.setBorderRight(CellStyle.BORDER_THIN);
        BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT.setBorderLeft(CellStyle.BORDER_THIN);

        BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT = getSimpleCellStyle(true, true, 0, 2, 1, 0);

        BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT = getSimpleCellStyle(true, true, 0, 2, 0, 1);

        BOLD_CENTER_MIDDLE_RIGHT_LEFT = getSimpleCellStyle(true, true, 0, 0, 1, 1);

        BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT.cloneStyleFrom(BOLD_CENTER_MIDDLE_RIGHT_LEFT);
        BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT.setBorderTop(CellStyle.BORDER_THICK);

        BOLD_DEFAULT_THICK_TOP_BOTTOM = getSimpleCellStyle(true, false, 2, 2, 0, 0);

        DEFAULT_CENTER_MIDDLE_RIGHT_LEFT = getSimpleCellStyle(false, true, 0, 0, 1, 1);

        DEFAULT_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        DEFAULT_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT.setBorderTop(CellStyle.BORDER_THICK);

        DEFAULT_CENTER_THICK_BOTTOM_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_CENTER_THICK_BOTTOM_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        DEFAULT_CENTER_THICK_BOTTOM_MIDDLE_RIGHT_LEFT.setBorderBottom(CellStyle.BORDER_THICK);

        DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT = getSimpleCellStyle(false, true, 0, 0, 1, 1);

        DEFAULT_DEFAULT_THICK_TOP_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_DEFAULT_THICK_TOP_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        DEFAULT_DEFAULT_THICK_TOP_MIDDLE_RIGHT_LEFT.setBorderTop(CellStyle.BORDER_THICK);

        DEFAULT_DEFAULT_THICK_BOTTOM_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_DEFAULT_THICK_BOTTOM_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        DEFAULT_DEFAULT_THICK_BOTTOM_MIDDLE_RIGHT_LEFT.setBorderBottom(CellStyle.BORDER_THICK);

        DEFAULT_THICK_RIGHT_LEFT_BG_GRAY = getSimpleCellStyle(false, false, 0, 0, 2, 2, IndexedColors.GREY_50_PERCENT.getIndex());
    }

    /**
     * get simple cell style
     * default : 맑은 고딕, font size 11, vertical center
     *
     * @param isBold       is font bold
     * @param isCenter     is alignment center
     * @param topBorder    top border. (2) thick, (1) thin, (others) default
     * @param bottomBorder bottom border. (2) thick, (1) thin, (others) default
     * @param leftBorder   left border. (2) thick, (1) thin, (others) default
     * @param rightBorder  right border. (2) thick, (1) thin, (others) default
     * @param foreColor    foreground color. 0 or negative is none.
     * @return cellStyle
     */
    public CellStyle getSimpleCellStyle(boolean isBold, boolean isCenter, int topBorder, int bottomBorder, int leftBorder, int rightBorder, short foreColor) {
        CellStyle cellStyle = this.workbook.createCellStyle();

        if (isBold) {
            cellStyle.setFont(getBoldFont(this.workbook));
        } else {
            cellStyle.setFont(getDefaultFont(this.workbook));
        }

        if (isCenter) {
            cellStyle.setAlignment(CellStyle.ALIGN_CENTER);
            cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        } else {
            cellStyle.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
        }

        if (topBorder == 2) {
            cellStyle.setBorderTop(CellStyle.BORDER_THICK);
        } else if (topBorder == 1) {
            cellStyle.setBorderTop(cellStyle.BORDER_THIN);
        }

        if (bottomBorder == 2) {
            cellStyle.setBorderBottom(CellStyle.BORDER_THICK);
        } else if (bottomBorder == 1) {
            cellStyle.setBorderBottom(cellStyle.BORDER_THIN);
        }

        if (rightBorder == 2) {
            cellStyle.setBorderRight(CellStyle.BORDER_THICK);
        } else if (rightBorder == 1) {
            cellStyle.setBorderRight(cellStyle.BORDER_THIN);
        }

        if (leftBorder == 2) {
            cellStyle.setBorderLeft(CellStyle.BORDER_THICK);
        } else if (leftBorder == 1) {
            cellStyle.setBorderLeft(cellStyle.BORDER_THIN);
        }

        if (foreColor > 0) {
            cellStyle.setFillForegroundColor(foreColor);
            cellStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        }

        return cellStyle;
    }

    /**
     * get simple cell style
     * default : 맑은 고딕, font size 11, vertical center
     *
     * @param isBold   is font bold
     * @param isCenter is alignment center
     * @return
     */
    public CellStyle getSimpleCellStyle(boolean isBold, boolean isCenter) {
        return getSimpleCellStyle(isBold, isCenter, 0, 0, 0, 0, (short) 0);
    }

    /**
     * get simple cell style
     * default : 맑은 고딕, font size 11, vertical center
     *
     * @param isBold       is font bold
     * @param isCenter     is alignment center
     * @param topBorder    top border. (2) thick, (1) thin, (others) default
     * @param bottomBorder bottom border. (2) thick, (1) thin, (others) default
     * @param leftBorder   left border. (2) thick, (1) thin, (others) default
     * @param rightBorder  right border. (2) thick, (1) thin, (others) default
     * @return cellStyle
     */
    public CellStyle getSimpleCellStyle(boolean isBold, boolean isCenter, int topBorder, int bottomBorder, int leftBorder, int rightBorder) {
        return getSimpleCellStyle(isBold, isCenter, topBorder, bottomBorder, leftBorder, rightBorder, (short) 0);
    }

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
     * get font style that set only bold
     *
     * @param workbook workbook to set font style
     * @return Font
     */
    private static Font getBoldFont(Workbook workbook) {
        Font font = getDefaultFont(workbook);
        font.setBoldweight(Font.BOLDWEIGHT_BOLD);

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
    private static Font getDefaultFont(Workbook workbook) {
        Font defaultFont = workbook.createFont();

        defaultFont.setFontName(POIConstant.DEFAULT_FONT_NAME);

        short defaultFontHeight = POIConstant.DEFAULT_FONT_HEIGHT * POIConstant.FONT_SIZE_CONSTANT;
        defaultFont.setFontHeight(defaultFontHeight);

        return defaultFont;
    }
}