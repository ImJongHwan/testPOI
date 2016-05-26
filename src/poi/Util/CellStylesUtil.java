package poi.Util;

import org.apache.poi.ss.format.CellFormat;
import org.apache.poi.ss.format.CellFormatType;
import org.apache.poi.ss.usermodel.*;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STCellType;
import poi.Constant;

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
    public CellStyle BOOLEAN_DEFAULT_MIDDLE_RIGHT_LEFT;

    private Workbook workbook = null;

    public CellStylesUtil(Workbook workbook) {
        this.workbook = workbook;
        init(workbook);
    }

    private void init(Workbook workbook) {
        BOLD_CENTER_THICK_TOP_BOTTOM = getBoldCenterStyle(workbook);
        BOLD_CENTER_THICK_TOP_BOTTOM.setBorderTop(CellStyle.BORDER_THICK);
        BOLD_CENTER_THICK_TOP_BOTTOM.setBorderBottom(CellStyle.BORDER_THICK);

        BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT.cloneStyleFrom(BOLD_CENTER_THICK_TOP_BOTTOM);
        BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT.setBorderRight(CellStyle.BORDER_THIN);
        BOLD_CENTER_THICK_TOP_BOTTOM_MIDDLE_RIGHT_LEFT.setBorderLeft(CellStyle.BORDER_THIN);

        BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT = getBoldCenterStyle(workbook);
        BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT.setBorderBottom(CellStyle.BORDER_THICK);
        BOLD_CENTER_THICK_BOTTOM_MIDDLE_LEFT.setBorderLeft(CellStyle.BORDER_THIN);

        BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT = getBoldCenterStyle(workbook);
        BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT.setBorderBottom(CellStyle.BORDER_THICK);
        BOLD_CENTER_THICK_BOTTOM_MIDDLE_RIGHT.setBorderRight(CellStyle.BORDER_THIN);

        BOLD_CENTER_MIDDLE_RIGHT_LEFT = getBoldCenterStyle(workbook);
        BOLD_CENTER_MIDDLE_RIGHT_LEFT.setBorderRight(CellStyle.BORDER_THIN);
        BOLD_CENTER_MIDDLE_RIGHT_LEFT.setBorderLeft(CellStyle.BORDER_THIN);

        BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT.cloneStyleFrom(BOLD_CENTER_MIDDLE_RIGHT_LEFT);
        BOLD_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT.setBorderTop(CellStyle.BORDER_THICK);

        BOLD_DEFAULT_THICK_TOP_BOTTOM = getBoldStyle(workbook);
        BOLD_DEFAULT_THICK_TOP_BOTTOM.setBorderTop(CellStyle.BORDER_THICK);
        BOLD_DEFAULT_THICK_TOP_BOTTOM.setBorderBottom(CellStyle.BORDER_THICK);

        DEFAULT_CENTER_MIDDLE_RIGHT_LEFT = getDefaultCenterStyle(workbook);
        DEFAULT_CENTER_MIDDLE_RIGHT_LEFT.setBorderRight(CellStyle.BORDER_THIN);
        DEFAULT_CENTER_MIDDLE_RIGHT_LEFT.setBorderLeft(CellStyle.BORDER_THIN);

        DEFAULT_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        DEFAULT_CENTER_THICK_TOP_MIDDLE_RIGHT_LEFT.setBorderTop(CellStyle.BORDER_THICK);

        DEFAULT_CENTER_THICK_BOTTOM_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_CENTER_THICK_BOTTOM_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_CENTER_MIDDLE_RIGHT_LEFT);
        DEFAULT_CENTER_THICK_BOTTOM_MIDDLE_RIGHT_LEFT.setBorderBottom(CellStyle.BORDER_THICK);

        DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT.setBorderRight(CellStyle.BORDER_THIN);
        DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT.setBorderLeft(CellStyle.BORDER_THIN);

        BOOLEAN_DEFAULT_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        BOOLEAN_DEFAULT_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);

        DEFAULT_DEFAULT_THICK_TOP_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_DEFAULT_THICK_TOP_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        DEFAULT_DEFAULT_THICK_TOP_MIDDLE_RIGHT_LEFT.setBorderTop(CellStyle.BORDER_THICK);

        DEFAULT_DEFAULT_THICK_BOTTOM_MIDDLE_RIGHT_LEFT = workbook.createCellStyle();
        DEFAULT_DEFAULT_THICK_BOTTOM_MIDDLE_RIGHT_LEFT.cloneStyleFrom(DEFAULT_DEFAULT_MIDDLE_RIGHT_LEFT);
        DEFAULT_DEFAULT_THICK_BOTTOM_MIDDLE_RIGHT_LEFT.setBorderBottom(CellStyle.BORDER_THICK);

        DEFAULT_THICK_RIGHT_LEFT_BG_GRAY = workbook.createCellStyle();
        DEFAULT_THICK_RIGHT_LEFT_BG_GRAY.setBorderRight(CellStyle.BORDER_THIN);
        DEFAULT_THICK_RIGHT_LEFT_BG_GRAY.setBorderLeft(CellStyle.BORDER_THIN);
        DEFAULT_THICK_RIGHT_LEFT_BG_GRAY.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
        DEFAULT_THICK_RIGHT_LEFT_BG_GRAY.setFillPattern(CellStyle.SOLID_FOREGROUND);
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

    public static CellStyle getDefaultCenterStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setFont(setDefaultFont(workbook));

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

    public static CellStyle setDefaultStyle(Workbook workbook) {
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
    private static Font setDefaultFont(Workbook workbook) {
        Font defaultFont = workbook.createFont();

        defaultFont.setFontName(Constant.DEFAULT_FONT_NAME);

        short defaultFontHeight = Constant.DEFAULT_FONT_HEIGHT * Constant.FONT_SIZE_CONSTANT;
        defaultFont.setFontHeight(defaultFontHeight);

        return defaultFont;
    }

    /**
     * append Style Top Border
     *
     * @param cellStyle   cellStyle
     * @param borderStyle borderStyle, you can get a CellStyle class.
     */
    public static void appendTopBorder(CellStyle cellStyle, short borderStyle) {
        if (cellStyle != null) {
            cellStyle.setBorderTop(borderStyle);
        }
    }

    /**
     * append Style Bottom Border
     *
     * @param cellStyle   cellStyle
     * @param borderStyle borderStyle, you can get a CellStyle class.
     */
    public static void appendBottomBorder(CellStyle cellStyle, short borderStyle) {
        if (cellStyle != null) {
            cellStyle.setBorderBottom(borderStyle);
        }
    }

    /**
     * append Style Right Border
     *
     * @param cellStyle   cellStyle
     * @param borderStyle borderStyle, you can get a CellStyle class.
     */
    public static void appendRightBorder(CellStyle cellStyle, short borderStyle) {
        if (cellStyle != null) {
            cellStyle.setBorderRight(borderStyle);
        }
    }

    /**
     * append Style Left Border
     *
     * @param cellStyle   cellStyle
     * @param borderStyle borderStyle, you can get a CellStyle class.
     */
    public static void appendLeftBorder(CellStyle cellStyle, short borderStyle) {
        if (cellStyle != null) {
            cellStyle.setBorderLeft(borderStyle);
        }
    }
}