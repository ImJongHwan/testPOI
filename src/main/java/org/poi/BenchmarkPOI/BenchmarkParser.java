package org.poi.BenchmarkPOI;

import org.poi.Constant;
import org.poi.Util.FileUtil;
import org.poi.Util.StringUtil;

import java.io.File;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Hwan on 2016-05-27.
 */
public class BenchmarkParser {
    static public String BENCHMARK_TC_PATH = "C:\\gitProjects\\simpleURLParser\\benchmark\\benchmark_tc\\";
    final static public String BENCHMARK_RES_PATH = "C:\\gitProjects\\simpleURLParser\\benchmark\\benchmark_res\\";
    final static public String BENCHMARK_SCORE_PATH = "C:\\gitProjects\\simpleURLParser\\benchmark\\benchmark_score\\";

    public static final String BENCHMARK_HTML_TAIL = ".html";
    public static final String BENCHMARK_QUESTION_TAIL = "?";
    public static final String BENCHMARK_CONTAIN_TEST = "BenchmarkTest";
    public static final String BENCHMARK_START_STRING =  "/benchmark/";

    public static final String BENCHMARK_CONTAIN_STRING = "BenchmarkTest";

    public static final String BENCHMARK_TEST_PREFIX = "benchmark_";
    public static final String BENCHMARK_TEST_CRAWLED = "crawled_";

    /**
     * get failed true positive List
     *
     * @param targetFilePath target file path
     * @param vulnerability vulnerability
     * @return failed tp list
     */
    protected static List<String> getFailedTPList(String targetFilePath, String vulnerability){
        String expectedTPPath = BENCHMARK_TC_PATH + vulnerability + Constant.TRUE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(parseList(FileUtil.readFile(new File(targetFilePath))), FileUtil.readFile(new File(expectedTPPath)));
    }

    /**
     * get failed false positive list
     *
     * @param targetFilePath target file path
     * @param vulnerability vulnerability
     * @return failed fp lsit
     */
    protected static List<String> getFailedFPList(String targetFilePath, String vulnerability){
        String expectedFPPath = BENCHMARK_TC_PATH + vulnerability + Constant.FALSE_POSITIVE_POSTFIX + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(parseList(FileUtil.readFile(new File(targetFilePath))), FileUtil.readFile(new File(expectedFPPath)));
    }

    /**
     * get failed crawling list
     *
     * @param targetFile target file
     * @return failed crawled list
     */
    protected static List<String> getFailedCrawlingList(File targetFile){
        List<String> crawledList = parseList(FileUtil.readFile(targetFile));

        String expectedCrawlingFilePath = BENCHMARK_TC_PATH + Constant.ALL_TC_FILE_NAME + Constant.TC_FILE_EXTENSION;
        return StringUtil.getComplementList(crawledList, FileUtil.readFile(expectedCrawlingFilePath));
    }

    /**
     * get failed crawling list
     *
     * @param targetFilePath target file path
     * @param originFilePath origin file path
     * @return failed crawled list
     */
    protected static List<String> getFailedCrawlingList(String targetFilePath, String originFilePath){
        File targetFile = new File(targetFilePath);
        File originFile = new File(originFilePath);

        return StringUtil.getComplementList(parseList(FileUtil.readFile(targetFile)), FileUtil.readFile(originFile));
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
            System.err.println("Can't remove tails since targetList is null");
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
