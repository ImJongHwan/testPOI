package poi;

/**
 * Created by Hwan on 2016-05-10.
 */
public class Constant {
    public static final String BENCHMARK_NAME = "ZAP_OWASPBenchmark_Results_";
    public static final String WAVSEP_NAME = "ZAP_WAVSEP_Results_";

    public static final String FILE_EXTENSION = ".xlsx";

    public enum BenchmarkSheets{
        total               ("Total"),
        cmdi                ("Command Injection"),
        crypto              ("Weak Encryption Algorithm"),
        hash                ("Weak Hash Algorithm"),
        ldapi               ("LDAP Injection"),
        pathtraver          ("Path Traversal"),
        securecookie        ("Insecure Cookie"),
        sqli                ("SQL Injection"),
        trustbound          ("Trust Boundary Violation"),
        weakrand            ("Weak Random Number"),
        xpathi              ("XPath Injection"),
        xss                 ("Cross-Site Scripting");

        private String sheetName;

        BenchmarkSheets(String sheetName){
            this.sheetName = sheetName;
        }

        public String getSheetName() {
            return sheetName;
        }
    }
}
