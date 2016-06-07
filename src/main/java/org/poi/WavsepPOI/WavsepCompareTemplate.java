package org.poi.WavsepPOI;

import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.FileUtil;

import java.io.IOException;

/**
 * Created by Hwan on 2016-06-07.
 */
public class WavsepCompareTemplate {

    private TcWorkbook beforeWorkbook = null;
    private TcWorkbook afterWorkbook = null;
    private TcWorkbook compareWorkbook = null;

    public WavsepCompareTemplate(TcWorkbook beforeWorkbook, TcWorkbook afterWorkbook){
        this.beforeWorkbook = beforeWorkbook;
        this.afterWorkbook = afterWorkbook;
        this.compareWorkbook = new TcWorkbook();

        init();
    }

    public WavsepCompareTemplate(String beforePath, String afterPath) throws IOException {
        this(new TcWorkbook(FileUtil.readExcelFile(beforePath)), new TcWorkbook(FileUtil.readExcelFile(afterPath)));
    }

    private void init() {
        setCompareStyle();
        setCompareWidth();
        setCompareValues();
    }

    /**
     * set compare workbook style
     */
    private void setCompareStyle() {

    }

    /**
     * set compare workbook width
     */
    private void setCompareWidth(){

    }

    /**
     * set compare workbook values
     */
    private void setCompareValues(){

    }

    /**
     * Write current workbook state in a file
     *
     * @param outputDir output directory path
     * @param fileName file name
     */
    public void writeCompareWorkbook(String outputDir, String fileName){

    }
}
