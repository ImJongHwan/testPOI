package poi;

/**
 * Created by Hwan on 2016-05-10.
 */
public class Constant {
    public static final String BENCHMARK_NAME = "ZAP_OWASPBenchmark_Results_";
    public static final String WAVSEP_NAME = "ZAP_WAVSEP_Results_";

    public static final String FILE_EXTENSION = ".xlsx";

    public static final String MAX_COLUMN_STRING = "XFD";
    public static final int MAX_ROW = 1048576;

    public static final String RES_OUTPUT_PATH = "C:\\scalaProjects\\testPOI\\results\\";

    public static final int BENCHMARK_TEST_CASE_CELL_SIZE = 25;
    public static final int BOOLEAN_CELL_SIZE = 15;
    public static final int WAVSEP_TEST_CASE_CELL_SIZE = 50;

    public static final int ALPHABET_SIZE = 26;
    public static final int COLUMN_SIZE_CONSTANT = 256;
    public static final short FONT_SIZE_CONSTANT = 20;

    public static final String DEFAULT_FONT_NAME = "맑은 고딕";
    public static final short DEFAULT_FONT_HEIGHT = 11;

    public enum BenchmarkSheets {
        total("Total"),
        cmdi("Command Injection"),
        crypto("Weak Encryption Algorithm"),
        hash("Weak Hash Algorithm"),
        ldapi("LDAP Injection"),
        pathtraver("Path Traversal"),
        securecookie("Insecure Cookie"),
        sqli("SQL Injection"),
        trustbound("Trust Boundary Violation"),
        weakrand("Weak Random Number"),
        xpathi("XPath Injection"),
        xss("Cross-Site Scripting");

        private String sheetName;

        BenchmarkSheets(String sheetName) {
            this.sheetName = sheetName;
        }

        public String getSheetName() {
            return sheetName;
        }
    }

    /**
     * Convert to column index to column alphabet
     *
     * @param columnIndex XSSFCell column index
     * @return columnAlphabet
     */
    public static String convertColumnIndexToAlphabet(int columnIndex) {
        String columnAlphabet = null;

        int tempIndex = 0;
        int resIndex = columnIndex;

        while (resIndex > 0) {
            tempIndex = (resIndex % ALPHABET_SIZE) + 'A';
            columnAlphabet = String.valueOf(Character.toChars(tempIndex)) + columnAlphabet;
            resIndex = resIndex / ALPHABET_SIZE;
        }

        return columnAlphabet;
    }

    /**
     * convert ColumnAlphabet in real excel to Cell Index
     *
     * @param columnAlphabet columnAlphabet in real excel
     * @return cell index
     */
    public static int convertColumnAlphabetToIndex(String columnAlphabet) {
        columnAlphabet = columnAlphabet.toUpperCase();

        int cellIndex = 0;

        for (int i = 0; i < columnAlphabet.length(); i++) {
            cellIndex += (columnAlphabet.charAt(i) - 'A') * Math.pow(ALPHABET_SIZE, columnAlphabet.length() - i - 1);
        }

        return cellIndex;
    }

    /**
     * convert ColumnAlphabet in real excel to Cell Index
     *
     * @param columnAlphabet columnAlphabet in real excel
     * @return cell index
     */
    public static int convertColumnAlphabetToIndex(char columnAlphabet) {
        String columnString = String.valueOf(columnAlphabet);

        return convertColumnAlphabetToIndex(columnString);
    }
}
