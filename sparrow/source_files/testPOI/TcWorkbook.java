package poi.TestCasePOI;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import poi.Constant;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by Hwan on 2016-05-10.
 */
public class TcWorkbook {

    private XSSFWorkbook workbook = null;
    private int newFileNum = 1;

    private Date workbookCreatedTime = null;

    private Map<String, TcSheet> benchmarkSheetMap = new HashMap<>();

    public TcWorkbook(){
        this.workbook = new XSSFWorkbook();
        init();
    }

    public TcWorkbook(XSSFWorkbook workbook){
        this.workbook = workbook;
        init();
    }

    /**
     * Set created workbook time.
     */
    private void init(){
//        SimpleDateFormat sdf = new SimpleDateFormat("yyMMddhhmm");
        this.workbookCreatedTime = new Date();
    }

    /**
     * Create and write a workbook current state.
     *
     * @param  tcName test cases name. It will be file main name.
     * @return made file Absolute Path. If creating file is failed, null return.
     */
    public String writeWorkbook(String tcName){
        try {

            SimpleDateFormat sdf = new SimpleDateFormat("yyMMddhhmm");
            String createdTime = sdf.format(workbookCreatedTime);

            String workbookFileName = tcName + createdTime + Constant.FILE_EXTENSION;
            File workbookFile = new File(workbookFileName);

            while(workbookFile.exists()) {
                workbookFileName = tcName + createdTime +" (" + this.newFileNum + ")" + Constant.FILE_EXTENSION;
                workbookFile = new File(workbookFileName);
                this.newFileNum++;
            }

            FileOutputStream out = new FileOutputStream(workbookFile);
            this.workbook.write(out);

            out.close();

            return workbookFile.getAbsolutePath();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("Creating a workbook file Error - Creating a test cases workbook file is failed.");
        } catch (IOException e){
            e.printStackTrace();
            System.out.println("Writing a workbook Error - Writing a test cases workbook is failed.");
        }

        return null;

    }

    /**
     * Create vulnerability sheet template.
     *
     * @param sheetName vulnerabilityName. Also it is sheet name.
     * @return vulnerability sheet.
     */
    public XSSFSheet getSheet(String sheetName) {

        return getTcSheet(sheetName).getSheet();

    }

    public TcSheet getTcSheet(String sheetName){
        if(sheetName == null){
            System.err.println("VulnerabilityName Error - VulnerabilityName is null. Please input vulnerabilityName.");
            return null;
        }

        if(benchmarkSheetMap.containsKey(sheetName)){
            return benchmarkSheetMap.get(sheetName);
        }

        XSSFSheet vulnerabilitySheet = workbook.createSheet(sheetName);
        TcSheet tcSheet = new TcSheet(workbook, vulnerabilitySheet);
        benchmarkSheetMap.put(sheetName, tcSheet);

        return tcSheet;
    }

    /**
     * Get workbook created time
     *
     * @return workbook createdTime
     */
    public Date getWorkbookCreatedTime() {
        return workbookCreatedTime;
    }

    /**
     * Get XSSFWorkbook instant.
     *
     * @return XSSFWorkbook instant
     */
    public XSSFWorkbook getWorkbook() {
        return workbook;
    }
}