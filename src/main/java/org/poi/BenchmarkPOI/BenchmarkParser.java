package org.poi.BenchmarkPOI;

import org.poi.Constant;
import org.poi.Util.FileUtil;
import org.poi.Util.StringUtil;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Hwan on 2016-05-27.
 */
public class BenchmarkParser {
    public static String BENCHMARK_TC_PATH = "BenchmarkTC/";

    public static final String BENCHMARK_HTML_TAIL = ".html";
    public static final String BENCHMARK_QUESTION_TAIL = "?";
    public static final String BENCHMARK_CONTAIN_TEST = "BenchmarkTest";
    public static final String BENCHMARK_START_STRING =  "/benchmark/";

    public static final String BENCHMARK_TEST_PREFIX = "benchmark_";
    public static final String BENCHMARK_TEST_CRAWLED = "crawled_";

    /**
     * get failed true positive List
     *
     * @param targetFilePath target file path
     * @param vulnerability vulnerability
     * @return failed tp list
     */
    protected static List<String> getFailedTPList(String targetFilePath, String vulnerability) throws IOException {
        String expectedTPPath = BENCHMARK_TC_PATH + vulnerability + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(parseList(FileUtil.readFile(new File(targetFilePath))), FileUtil.readResourceFile(BenchmarkParser.class, expectedTPPath));
    }



    /**
     * get failed false positive list
     *
     * @param targetFilePath target file path
     * @param vulnerability vulnerability
     * @return failed fp lsit
     */
    protected static List<String> getFailedFPList(String targetFilePath, String vulnerability) throws IOException {
        String expectedFPPath = BENCHMARK_TC_PATH + vulnerability + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(parseList(FileUtil.readFile(new File(targetFilePath))), FileUtil.readResourceFile(BenchmarkParser.class, expectedFPPath));
    }


    /**
     * get failed crawling list
     *
     * @param targetFile target file
     * @return failed crawled list
     */
    protected static List<String> getFailedCrawlingList(File targetFile, String vulnerability) throws IOException {
        List<String> pureCrawledList = FileUtil.readFile(targetFile);

        List<String> expectedCrawlingList = new ArrayList<>();
        List<String> tempList;
        if((tempList = FileUtil.readResourceFile(BenchmarkParser.class, BENCHMARK_TC_PATH + vulnerability + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION)) != null){
            expectedCrawlingList.addAll(tempList);
        }
        if((tempList = FileUtil.readResourceFile(BenchmarkParser.class, BENCHMARK_TC_PATH + vulnerability + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION)) != null){
            expectedCrawlingList.addAll(tempList);
        }

        return getFailedCrawlingList(pureCrawledList, expectedCrawlingList);
    }

    /**
     * get failed crawling list
     *
     * @param pureCrawledList pure crawled list, not parse list
     * @param expectedCrawlingList expected crawling list
     * @return failed crawling list
     */
    public static List<String> getFailedCrawlingList(List<String> pureCrawledList, List<String> expectedCrawlingList){
        return StringUtil.getComplementList(parseList(pureCrawledList), expectedCrawlingList);
    }

    /**
     * get false negative list
     *
     * @param pureScannedList pure scanned list, not parsing list
     * @param expectedScannedList expected scanned list
     * @return false negative list
     */
    public static List<String> getFalseNegativeList(List<String> pureScannedList, List<String> expectedScannedList){
        return StringUtil.getComplementList(parseList(pureScannedList), expectedScannedList);
    }

    /**
     * get false positive list
     *
     * @param pureScannedList pure scanned list, not parsing list
     * @param expectedScannedList expected scanned list
     * @return false positive list
     */
    public static List<String> getTrueNegativeList(List<String> pureScannedList, List<String> expectedScannedList){
        return StringUtil.getComplementList(parseList(pureScannedList), expectedScannedList);
    }

    /**
     * parse a url list
     * Benchmark alerts line form ex ) Command Injection : https://localhost:8443/benchmark/BenchmarkTest00006.html
     * > after parse : BenchmarkTest00006
     * Benchmark Crawled line form ex ) https://localhost:8443/benchmark/BenchmarkTest00006.html
     * > after parse : BenchmarkTest00006
     *
     * @param targetList target list
     * @return parsed list
     */
    private static List<String> parseList(List<String> targetList){
        List<String> mustContain = new ArrayList<>();
        mustContain.add(BENCHMARK_CONTAIN_TEST);
        List<String> parsingListContainedTail = StringUtil.parseList(targetList, BENCHMARK_START_STRING, mustContain);

        List<String> tailList = new ArrayList<>();

        tailList.add(BENCHMARK_HTML_TAIL);
        tailList.add(BENCHMARK_QUESTION_TAIL);

        return removeTails(parsingListContainedTail, tailList);
    }

    /**
     * get remove Tail list
     *
     * @param targetList target list
     * @param tailList tail list
     * @return removeTails list
     */
    private static List<String> removeTails(List<String> targetList, List<String> tailList){
        if(targetList == null){
            return null;
        }

        if(tailList == null || tailList.isEmpty()){
            return targetList;
        }

        List<String> removeTailList = new ArrayList<>();
        boolean hasPostfix;
        for(String target : targetList){
            hasPostfix = false;
            for(String tail : tailList){

                if(target.contains(tail)){
                    removeTailList.add(target.substring(0, target.indexOf(tail)));
                    hasPostfix = true;
                    break;
                }
            }
            if(!hasPostfix){
                removeTailList.add(target);
            }
        }
        return removeTailList;
    }

    /**
     * set Benchmark Tc path
     *
     * @param benchmarkTcPath benchmark tc path
     */
    public static void setBenchmarkTcPath(String benchmarkTcPath) {
        BENCHMARK_TC_PATH = benchmarkTcPath;
    }
}
