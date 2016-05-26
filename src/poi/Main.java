package poi;

import poi.BenchmarkPOI.BenchmarkTemplate;
import poi.WavsepPOI.WavsepTemplate;

public class Main {

    public static void main(String[] args) {
//        BenchmarkTemplate bt = new BenchmarkTemplate();
//        bt.writeBenchmark();

        WavsepTemplate wt = new WavsepTemplate();
        wt.writeWavsep();
    }
}
