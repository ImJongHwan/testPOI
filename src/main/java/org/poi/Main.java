package org.poi;

import org.poi.BenchmarkPOI.BenchmarkTemplate;
import org.poi.TestCasePOI.TcWorkbook;
import org.poi.Util.FileUtil;
import org.poi.WavsepPOI.WavsepTemplate;

import java.io.IOException;

public class Main {

    public static void main(String[] args) {
        if(args.length < 1) {
            return;
        } else {
            if(args[0].equals("-b")) {
                BenchmarkTemplate bt = new BenchmarkTemplate();
                bt.writeFailedList("C:\\gitProjects\\zap\\results\\160609153221_benchmarktest");
                bt.writeBenchmark();
            } else if (args[0].equals("-w")) {
                WavsepTemplate wt = new WavsepTemplate();
                wt.writeFailedList("C:\\gitProjects\\zap\\results\\160602140051_wavseptest");
                wt.writeWavsep();
            } else if (args[0].equals("-a")){
                BenchmarkTemplate bt = new BenchmarkTemplate();
                bt.writeBenchmark();
                WavsepTemplate wt = new WavsepTemplate();
                wt.writeWavsep();
            } else if (args[0].equals("-rb")){
                try {
                    TcWorkbook bt = new TcWorkbook(FileUtil.readExcelFile("C:\\gitProjects\\zap\\results\\160602145110_wavseptest\\Zap_Wavsep_Results_160602145110.xlsx"));
                    System.out.println(bt.getWorkbookCreatedTime());
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else {
                System.out.println("Check a argument.");
            }
        }
    }
}
