package org.poi.Util;

import java.util.ArrayList;
import java.util.List;

/**
 * Created by Hwan on 2016-05-27.
 */
public class StringUtil {
    /**
     * get a complement list that (origin list - target list)
     *
     * @param targets target list
     * @param origins origin list
     * @return complement list
     */
    public static List<String> getComplementList(List<String> targets, List<String> origins){
        if(targets == null || origins == null){
            System.err.println("Comparing Error - list is null : StringUtil Class");
            return null;
        }

        if(targets.isEmpty() || origins.isEmpty()){
            System.out.println("StringUtil Class : Cannot compare targets and origins since target or origin list is empty");
            return null;
        }

        List<String> compareList = new ArrayList<>();

        for(String origin : origins){
            if(!targets.contains(origin)){
                compareList.add(origin);
            }
        }
        return compareList;
    }

    /**
     * get parse list from startParse to endParse contained certain strings
     * ex) {startParse} contents + certain strings {endParse}
     *     You can get a contents + certain strings list
     *
     * @param lineList list to parse
     * @param startParse start parsing string
     * @param endParse end parsing string
     * @param mustContainStrings must contain strings
     * @return string list
     */
    public static List<String> parseList(List<String> lineList, String startParse, String endParse, List<String> mustContainStrings){
        if (lineList == null || lineList.isEmpty()) {
            return null;
        }

        ArrayList<String> parseList = new ArrayList<>();

        for (String line : lineList) {
            String trimLine = line.trim();

            int startIndex = trimLine.indexOf(startParse) + startParse.length();
            String parseString;
            if (startIndex > 0) {

                if (endParse != null) {

                    int endIndex = trimLine.indexOf(endParse);

                    if (endIndex > 0 && endIndex > startIndex) {
                        parseString = trimLine.substring(startIndex, endIndex + endParse.length());
                    } else {
                        continue;
                    }
                }else {
                    parseString = trimLine.substring(startIndex);
                }

                if (mustContainStrings == null || mustContainStrings.isEmpty()) {
                    parseList.add(parseString);
                } else {
                    int i;
                    for (i = 0; i < mustContainStrings.size(); i++) {
                        if (!parseString.contains(mustContainStrings.get(i))) {
                            break;
                        }
                    }
                    if (i == mustContainStrings.size()) {
                        parseList.add(parseString);
                    }
                }
            }
        }
        return parseList;
    }

    /**
     * get parse list from startParse to end line contained certain strings
     * ex) {startParse} contents + certain strings
     *     You can get a contents + certain strings list
     *
     * @param lineList list to parse
     * @param startParse start parsing string
     * @param mustContainStrings must contain strings
     * @return string list
     */
    public static List<String> parseList(List<String> lineList, String startParse, List<String> mustContainStrings){
        return parseList(lineList, startParse, null, mustContainStrings);
    }

    /**
     * get parse list from startParse to end line
     * ex) {startParse} contents
     *     You can get a contents  list
     *
     * @param lineList list to parse
     * @param startParse start parsing string
     * @return string list
     */
    public static List<String> parseList(List<String> lineList, String startParse){
        return parseList(lineList, startParse, null, null);
    }

    /**
     * Get parse list from startParse to endParse
     * ex) startParse (content) endParse. We can get content list
     *
     * @param lineList   list to parse
     * @param startParse start parse string
     * @param endParse   end parse string
     * @return parsed string list
     */
    public static List<String> parseList(List<String> lineList, String startParse, String endParse) {
        return parseList(lineList, startParse, endParse, null);
    }
}
