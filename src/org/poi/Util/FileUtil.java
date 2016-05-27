package org.poi.Util;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Hwan on 2016-05-26.
 */
public class FileUtil {
    /**
     * read a file and put lines in string list
     * @param path file path
     * @return line string list
     */
    public static List<String> readFile(String path){
        File file = new File(path);

        if(!file.exists()){
            System.out.println("File don't exist - " + path);
            return null;
        }

        try {
            BufferedReader br = new BufferedReader(new FileReader(file));
            List<String> lines = new ArrayList<>();
            String line;

            while((line = br.readLine()) != null){
                lines.add(line);
            }

            br.close();

            return lines;
        } catch (IOException e) {
            System.out.println("Reading File is failed");
            e.printStackTrace();
        }
        return null;
    }

    /**
     * read a file and put lines in string list
     * @param file file
     * @return line string list
     */
    public static List<String> readFile(File file){
        if(!file.exists()){
            System.out.println("File don't exist - " + file.getAbsolutePath());
            return null;
        }

        try {
            BufferedReader br = new BufferedReader(new FileReader(file));
            List<String> lines = new ArrayList<>();
            String line;

            while((line = br.readLine()) != null){
                lines.add(line);
            }

            br.close();

            return lines;
        } catch (IOException e) {
            System.out.println("Reading File is failed");
            e.printStackTrace();
        }
        return null;
    }
}
