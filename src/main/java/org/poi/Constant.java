package org.poi;

import java.io.File;

/**
 * Created by Hwan on 2016-05-27.
 */
public class Constant {

    public static final String TRUE_POSITIVE_POSTFIX = "_tp";
    public static final String FALSE_POSITIVE_POSTFIX = "_fp";
    public static final String EXPERIMENTAL_POSTFIX = "_ex";
    public static final String TC_FILE_EXTENSION = ".txt";
    public static final String ALL_TC_FILE_NAME = "all";

    public static final String TEST_SET_FAILED_TRUE_POSITIVE = "failedTruePositive";
    public static final String TEST_SET_FAILED_FALSE_POSITIVE = "failedFalsePositive";
    public static final String TEST_SET_FAILED_EXPERIMENTAL = "failedExperimental";
    public static final String TEST_SET_FAILED_CRAWLING = "failedCrawling";

    public enum BenchmarkSheets {
        total("Total", -1),
        cmdi("Command Injection", 78),
        rxss("Cross-Site Scripting", 79),
        securecookie("Insecure Cookie", 614),
        ldapi("LDAP Injection", 90),
        pathtraver("Path Traversal", 22),
        sqli("SQL Injection", 89),
        trustbound("Trust Boundary Violation", 501),
        crypto("Weak Encryption Algorithm", 327),
        hash("Weak Hash Algorithm", 328),
        weakrand("Weak Random Number", 330),
        xpathi("XPath Injection", 643);


        private String sheetName;
        private int cwe;

        BenchmarkSheets(String sheetName, int cwe) {
            this.sheetName = sheetName;
            this.cwe = cwe;
        }

        public String getSheetName() {
            return sheetName;
        }

        public int getCwe() {
            return cwe;
        }
    }

    public enum WavsepSheets {
        total("Total"),
        lfi("Path Traversal(LFI)"),
        rfi("Remote File Inclusion"),
        rxss("Cross Site Scripting(Reflected)"),
        sqli("SQL Injection(SQLI)"),
        redirect("Unvalidated Redirect");

        private String sheetName;

        WavsepSheets(String sheetName) { this.sheetName = sheetName; }

        public String getSheetName() {
            return sheetName;
        }
    }

    public static final String TEST_CASES_COLUMN = "B";

    public enum WavsepTypeColumn {
        typeTruePositive("C"),
        typeFalsePositive("D"),
        typeExperimental("E");

        private String column;

        WavsepTypeColumn(String column){
            this.column = column;
        }

        public String getColumn() {
            return column;
        }
    }

    public enum WavsepCrawlColumn {
        urlCrawlTrue("F"),
        urlCrawlFalse("G");
        private String column;

        WavsepCrawlColumn(String column){
            this.column = column;
        }

        public String getColumn() {
            return column;
        }
    }

    public enum WavsepDetectedColumn {
        detectedTruePositiveTrue("H"),
        detectedTruePositiveFalse("I"),
        detectedFalsePositiveTrue("J"),
        detectedFalsePositiveFalse("K"),
        detectedExperimentalTrue("L"),
        detectedExperimentalFalse("M");

        private String column;

        WavsepDetectedColumn(String column){
            this.column = column;
        }

        public String getColumn() {
            return column;
        }
    }
}
