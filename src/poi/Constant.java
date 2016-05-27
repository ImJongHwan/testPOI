package poi;

/**
 * Created by Hwan on 2016-05-10.
 */
public class Constant {
    public static final String BENCHMARK_NAME = "ZAP_OWASPBenchmark_Results_";
    public static final String WAVSEP_NAME = "ZAP_WAVSEP_Results_";

    public static final String TC_BENCHMARK_PATH = "C:\\gitProjects\\simpleURLParser\\benchmark\\benchmark_tc\\";
    public static final String TC_WAVSEP_PATH = "C:\\gitProjects\\simpleURLParser\\wavsep\\wavsep_tc\\";

    public static final String TRUE_POSITIVE_POSTFIX = "_tp";
    public static final String FALSE_POSITIVE_POSTFIX = "_fp";
    public static final String EXPERIMENTAL_POSTFIX = "_ex";
    public static final String TC_FILE_EXTENSION = ".txt";

    public static final String FILE_EXTENSION = ".xlsx";

    public static final String MAX_COLUMN_STRING = "XFD";
    public static final int MAX_ROW = 1048576;

    public static final String RES_OUTPUT_PATH = "C:\\scalaProjects\\testPOI\\results\\";

    public static final int BENCHMARK_TEST_CASE_CELL_WIDTH = 25;
    public static final int DEFAULT_COLUMN_WIDTH = 10;
    public static final int WAVSEP_TEST_CASE_CELL_WIDTH = 50;

    public static final int ALPHABET_SIZE = 26;
    public static final int COLUMN_SIZE_CONSTANT = 256;
    public static final short FONT_SIZE_CONSTANT = 20;

    public static final String DEFAULT_FONT_NAME = "맑은 고딕";
    public static final short DEFAULT_FONT_HEIGHT = 11;

    public enum BenchmarkSheets {
        total("Total", -1),
        cmdi("Command Injection", 78),
        xss("Cross-Site Scripting", 79),
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
        rfi("Remote File Inclusion(XSS via RFI)"),
        rxss("Cross Site Scripting(Reflected)"),
        sqli("SQL Injection(SQLI)"),
        redirect("Unvalidated Redirect");

        private String sheetName;

        WavsepSheets(String sheetName) { this.sheetName = sheetName; }

        public String getSheetName() {
            return sheetName;
        }
    }
}
