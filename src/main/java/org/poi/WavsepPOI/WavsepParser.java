package org.poi.WavsepPOI;

import org.poi.Constant;
import org.poi.Util.FileUtil;
import org.poi.Util.StringUtil;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Hwan on 2016-05-27.
 */
public class WavsepParser {
    public static final String WAVSEP_TEST_PREFIX = "wavsep_";
    public static final String WAVSEP_TEST_CRAWLED = "crawled_";

    private static final String WAVSEP_RES_PATH = "C:\\gitProjects\\simpleURLParser\\wavsep\\wavsep_res\\";
    private static final String WAVSEP_TC_PATH = "C:\\gitProjects\\simpleURLParser\\wavsep\\wavsep_tc\\";
    private static final String WAVSEP_SCORE_PATH = "C:\\gitProjects\\simpleURLParser\\wavsep\\wavsep_score\\";

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
    public static List<String> getFailedTcList(File targetFile, String vulnerability, String testSet) {
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
    private static List<String> getFailedTpList(List<String> targetList, String vulnerability) {
        String tpPath = WAVSEP_TC_PATH + vulnerability + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(targetList, FileUtil.readFile(tpPath));
    }

    /**
     * get a failed false positive list
     *
     * @param targetList    target list
     * @param vulnerability vulnerability name
     * @return failed false positive list
     */
    private static List<String> getFailedFpList(List<String> targetList, String vulnerability) {
        String tpPath = WAVSEP_TC_PATH + vulnerability + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(targetList, FileUtil.readFile(tpPath));
    }

    /**
     * get a failed experimental list
     *
     * @param targetList    target list
     * @param vulnerability vulnerability name
     * @return failed experimental list
     */
    private static List<String> getFailedExList(List<String> targetList, String vulnerability) {
        String tpPath = WAVSEP_TC_PATH + vulnerability + Constant.EXPERIMENTAL_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(targetList, FileUtil.readFile(tpPath));
    }

    /**
     * get a failed crawling list
     *
     * @param targetFile crawled target list
     * @return failed crawled list
     */
    public static List<String> getFailedCrawlingList(File targetFile) {
        List<String> crawledList = parseList(FileUtil.readFile(targetFile));

        String tpPath = WAVSEP_TC_PATH + Constant.ALL_TC_FILE_NAME + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(crawledList, FileUtil.readFile(tpPath));
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
