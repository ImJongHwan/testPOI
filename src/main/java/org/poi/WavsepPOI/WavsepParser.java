package org.poi.WavsepPOI;

import org.poi.Constant;
import org.poi.Main;
import org.poi.Util.FileUtil;
import org.poi.Util.StringUtil;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Hwan on 2016-05-27.
 */
public class WavsepParser {
    public static final String WAVSEP_TEST_PREFIX = "wavsep_";
    public static final String WAVSEP_TEST_CRAWLED = "crawled_";

    public static final String WAVSEP_TC_PATH = "WavsepTC/";

    private static final String WAVSEP_CONTAIN_CASE_STRING = "Case";
    private static final String WAVSEP_CONTAIN_URL_EXTENSION_JSP = ".jsp";

    private static final String WAVSEP_START_ACTIVE = "/active/";
    private static final String WAVSEP_END_JSP = ".jsp";
    private static final String WAVSEP_END_PARAMETER = "?";


    /**
     * get failed list
     * > target alerts file form - "wavsep_{vulnerability name}.txt"
     * > target crawled file for - "wavsep_{vulnerability name}_crawled.txt"
     *
     * @param targetFile    target file
     * @param vulnerability vulnerability name
     * @param testSet       test set - Constant has this constant string.
     * @return filed List
     */
    public static List<String> getFailedTcList(File targetFile, String vulnerability, String testSet) throws IOException{
        List<String> targetList = parseList(FileUtil.readFile(targetFile));

        switch (testSet) {
            case Constant.TEST_SET_FAILED_TRUE_POSITIVE:
                return getFailedTpList(targetList, vulnerability);
            case Constant.TEST_SET_FAILED_FALSE_POSITIVE:
                return getFailedFpList(targetList, vulnerability);
            case Constant.TEST_SET_FAILED_EXPERIMENTAL:
                return getFailedExList(targetList, vulnerability);
            default:
                System.err.println("WavsepParser : Test set is wrong - " + testSet);
                return null;
        }
    }

    /**
     * get a failed true positive list
     *
     * @param targetList    target list
     * @param vulnerability vulnerability name
     * @return failed true positive list
     */
    private static List<String> getFailedTpList(List<String> targetList, String vulnerability) throws IOException{
        String tpPath = WAVSEP_TC_PATH + vulnerability + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(targetList, FileUtil.readResourceFile(WavsepParser.class, tpPath));
    }

    /**
     * get a failed false positive list
     *
     * @param targetList    target list
     * @param vulnerability vulnerability name
     * @return failed false positive list
     */
    private static List<String> getFailedFpList(List<String> targetList, String vulnerability) throws IOException{
        String tpPath = WAVSEP_TC_PATH + vulnerability + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(targetList, FileUtil.readResourceFile(WavsepParser.class, tpPath));
    }

    /**
     * get a failed experimental list
     *
     * @param targetList    target list
     * @param vulnerability vulnerability name
     * @return failed experimental list
     */
    private static List<String> getFailedExList(List<String> targetList, String vulnerability) throws IOException {
        String tpPath = WAVSEP_TC_PATH + vulnerability + Constant.EXPERIMENTAL_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(targetList, FileUtil.readResourceFile(WavsepParser.class, tpPath));
    }

    /**
     * get a failed crawling list
     *
     * @param targetFile crawled target list
     * @return failed crawled list
     */
    public static List<String> getFailedCrawlingList(File targetFile, String vulnerability) throws IOException{
        List<String> crawledList = parseList(FileUtil.readFile(targetFile));

        List<String> expectedCrawlingList = new ArrayList<>();
        List<String> tempList;
        if ((tempList = FileUtil.readResourceFile(WavsepParser.class, WAVSEP_TC_PATH + vulnerability + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION)) != null) {
            expectedCrawlingList.addAll(tempList);
        }
        if((tempList = FileUtil.readResourceFile(WavsepParser.class, WAVSEP_TC_PATH + vulnerability + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION)) != null){
            expectedCrawlingList.addAll(tempList);
        }
        if((tempList = FileUtil.readResourceFile(WavsepParser.class, WAVSEP_TC_PATH + vulnerability + Constant.EXPERIMENTAL_POSTFIX + Constant.TC_FILE_EXTENSION)) != null) {
            expectedCrawlingList.addAll(tempList);
        }
        return StringUtil.getComplementList(crawledList, expectedCrawlingList);
    }

    /**
     * get failed crawling list
     *
     * @param crawledFilePath crawled file path
     * @param expectedCrawlingList expected crawling test cases list
     * @return failed crawling list
     */
    public static List<String> getFailedCrawlingList(String crawledFilePath, List<String> expectedCrawlingList){
        List<String> crawledList = null;
        if(crawledFilePath != null){
            crawledList = parseList(FileUtil.readFile(crawledFilePath));
        }
        return StringUtil.getComplementList(crawledList, expectedCrawlingList);
    }

    /**
     * get FALSE negative List
     *
     * @param scannedFilePath scanned file path
     * @param expectedScannedList expected scanned list
     * @return false negative list
     */
    public static List<String> getFalseNegativeList(String scannedFilePath, List<String> expectedScannedList){
        List<String> scannedList = null;
        if(scannedFilePath != null){
            scannedList = parseList(FileUtil.readFile(scannedFilePath));
        }
        return StringUtil.getComplementList(scannedList, expectedScannedList);
    }

    /**
     * get true negative list
     *
     * @param scannedFilePath scanned file path
     * @param expectedScannedList expected scanned list
     * @return true negative list
     */
    public static List<String> getTrueNegativeList(String scannedFilePath, List<String> expectedScannedList){
        List<String> scannedList = null;
        if(scannedFilePath != null){
            scannedList = parseList(FileUtil.readFile(scannedFilePath));
        }
        return StringUtil.getComplementList(scannedList, expectedScannedList);
    }

    /**
     * get experimental list
     *
     * @param scannedFilePath scanned file path
     * @param expectedScannedList expected scanned list
     * @return experimental list
     */
    public static List<String> getExperimentalList(String scannedFilePath, List<String> expectedScannedList){
        List<String> scannedList = null;
        if(scannedFilePath != null){
            scannedList = parseList(FileUtil.readFile(scannedFilePath));
        }
        return StringUtil.getComplementList(scannedList, expectedScannedList);
    }

    /**
     * parse a urls file
     * alert line form ex) Remote File Inclusion : http://localhost:8080/wavsep/active/Unvalidated-Redirect/Redirect-Detection-Evaluation-GET-302Redirect/Case01-Redirect-RedirectMethod-FilenameContext-Unrestricted-HttpURL-DefaultFullInput-AnyPathReq-Read.jsp?target=http%3A%2F%2Fwww.google.com%2F
     * after parse : Redirect-Detection-Evaluation-GET-302Redirect/Case01-Redirect-RedirectMethod-FilenameContext-Unrestricted-HttpURL-DefaultFullInput-AnyPathReq-Read
     * crawled line form ex) http://localhost:8080/wavsep/active/Unvalidated-Redirect/Redirect-Detection-Evaluation-POST-302Redirect/Case08-Redirect-RedirectMethod-FilenameContext-HttpInputValidation-HttpURL-DefaultPartialInput-PartialPathReq-Read.jsp
     * after parse : Redirect-Detection-Evaluation-POST-302Redirect/Case08-Redirect-RedirectMethod-FilenameContext-HttpInputValidation-HttpURL-DefaultPartialInput-PartialPathReq-Read
     *
     * @param targetList target list
     * @return parsed list
     */
    private static List<String> parseList(List<String> targetList) {
        List<String> mustContain = new ArrayList<>();
        mustContain.add(WAVSEP_CONTAIN_CASE_STRING);
        mustContain.add(WAVSEP_CONTAIN_URL_EXTENSION_JSP);
        List<String> firstParse = StringUtil.parseList(targetList, WAVSEP_START_ACTIVE, WAVSEP_END_JSP, mustContain);
        return StringUtil.parseList(firstParse, "/", mustContain);
    }
}
