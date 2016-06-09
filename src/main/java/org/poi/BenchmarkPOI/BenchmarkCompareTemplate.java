package org.poi.BenchmarkPOI;

import org.poi.Constant;
import org.poi.TestCasePOI.TcSheet;
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

    /**
     * BenchmarkCompareTemplate instant
     *
     * @param beforeWorkbook a before benchmark workbook
     * @param afterWorkbook a after benchmark workbook
     */
    public BenchmarkCompareTemplate(TcWorkbook beforeWorkbook, TcWorkbook afterWorkbook){
        this.beforeWorkbook = beforeWorkbook;
        this.afterWorkbook = afterWorkbook;
        this.compareWorkbook = new TcWorkbook();

        init();
    }

    /**
     * BenchmarkCompareTemplate instant
     *
     * @param beforePath a before benchmark workbook path
     * @param afterPath a after benchmark workbook path
     * @throws IOException
     */
    public BenchmarkCompareTemplate(String beforePath, String afterPath) throws IOException {
        this(new TcWorkbook(FileUtil.readExcelFile(beforePath)), new TcWorkbook(FileUtil.readExcelFile(afterPath)));
    }

    /** initialize Benchmark compare template */
    private void init(){
        for(Constant.BenchmarkSheets sheet : Constant.BenchmarkSheets.values()){
            if(sheet.toString().equals(Constant.BenchmarkSheets.total.toString())){
                setTotalSheet(this.compareWorkbook.getTcSheet(sheet.getSheetName()));
            } else {
                setVulnerabilitySheets(this.compareWorkbook.getTcSheet(sheet.getSheetName()));
            }
        }
    }

    /**
     * set Total sheet
     * @param totalSheet total sheet
     */
    private void setTotalSheet(TcSheet totalSheet){
        setCompareTotalStyles(totalSheet);
        setCompareTotalWidth(totalSheet);
        setCompareTotalValues(totalSheet);
    }

    /**
     * set compare total sheet styles
     *
     * @param sheet total sheet
     */
    private void setCompareTotalStyles(TcSheet sheet){

    }


    /**
     * set compare total sheet width
     *
     * @param sheet total sheet
     */
    private void setCompareTotalWidth(TcSheet sheet){

    }

    /**
     * set compare total sheet values
     *
     * @param sheet total sheet
     */
    private void setCompareTotalValues(TcSheet sheet){

    }

    /**
     * set a vulnerability sheet
     *
     * @param vulnerabilitySheet vulnerability sheet
     */
    private void setVulnerabilitySheets(TcSheet vulnerabilitySheet){
        setCompareVulnerabilityStyles(vulnerabilitySheet);
        setCompareVulnerabilityWidth(vulnerabilitySheet);
        setCompareVulnerabilityValues(vulnerabilitySheet);
    }

    /**
     * set a compare vulnerability sheet styles
     *
     * @param sheet vulnerability sheet
     */
    private void setCompareVulnerabilityStyles(TcSheet sheet){

    }

    /**
     * set a compare vulnerability sheet width
     *
     * @param sheet vulnerability sheet
     */
    private void setCompareVulnerabilityWidth(TcSheet sheet){

    }


    /**
     * set a compare vulnerability sheet values
     *
     * @param sheet vulnerability sheet
     */
    private void setCompareVulnerabilityValues(TcSheet sheet){

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
