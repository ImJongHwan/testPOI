package poi;

import poi.BenchmarkPOI.BenchmarkTemplate;
import poi.WavsepPOI.WavsepTemplate;

public class Main {

    public static void main(String[] args) {
        if(args.length < 1) {
            return;
        } else {
            if(args[0].equals("-b")) {
                BenchmarkTemplate bt = new BenchmarkTemplate();
                bt.writeBenchmark();
            } else if (args[0].equals("-w")) {
                WavsepTemplate wt = new WavsepTemplate();
                wt.writeWavsep();
            } else if (args[0].equals("-a")){
                BenchmarkTemplate bt = new BenchmarkTemplate();
                bt.writeBenchmark();
                WavsepTemplate wt = new WavsepTemplate();
                wt.writeWavsep();
            } else {
                System.out.println("Check a argument.");
            }
        }
    }
}
