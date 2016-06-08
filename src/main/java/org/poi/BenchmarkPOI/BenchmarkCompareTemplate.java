package org.poi.BenchmarkPOI;

import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.FileUtil;

import java.io.IOException;

/**
 * Created by Hwan on 2016-06-07.
 */
public class BenchmarkCompareTemplate {

    /** before TcWorkbook for compare */
    private TcWorkbook beforeWorkbook = null;
    /** after TcWorkbook for compare */
    private TcWorkbook afterWorkbook = null;
    /** compare TcWorkbook for compare */
    private TcWorkbook compareWorkbook = null;

    public BenchmarkCompareTemplate(TcWorkbook beforeWorkbook, TcWorkbook afterWorkbook){
        this.beforeWorkbook = beforeWorkbook;
        this.afterWorkbook = afterWorkbook;
        this.compareWorkbook = new TcWorkbook();

        init();
    }

    public BenchmarkCompareTemplate(String beforePath, String afterPath) throws IOException {
        this(new TcWorkbook(FileUtil.readExcelFile(beforePath)), new TcWorkbook(FileUtil.readExcelFile(afterPath)));
    }

    /** initialize Benchmark compare template */
    private void init(){
        setCompareStyle();
        setCompareWidth();
        setCompareValues();
    }

    /**
     * set compare workbook style
     */
    private void setCompareStyle(){

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
     * @return file absolute path
     */
    public String writeCompareWrokbook(String outputDir, String fileName){
        if(compareWorkbook != null){
            return compareWorkbook.writeWorkbook(outputDir, fileName);
        }
        return null;
    }
}
