package org.poi.Util;

import java.io.*;
import java.util.ArrayList;
import java.util.Collection;
import java.util.List;

/**
 * Created by Hwan on 2016-05-26.
 */
public class FileUtil {
    /**
     * read a file and put lines in string list
     *
     * @param path file path
     * @return line string list
     */
    public static List<String> readFile(String path) {
        File file = new File(path);

        if (!file.exists()) {
            System.out.println("File don't exist - " + path);
            return null;
        }

        try {
            BufferedReader br = new BufferedReader(new FileReader(file));
            List<String> lines = new ArrayList<>();
            String line;

            while ((line = br.readLine()) != null) {
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
     *
     * @param file file
     * @return line string list
     */
    public static List<String> readFile(File file) {
        if (!file.exists()) {
            System.out.println("File don't exist - " + file.getAbsolutePath());
            return null;
        }

        try {
            BufferedReader br = new BufferedReader(new FileReader(file));
            List<String> lines = new ArrayList<>();
            String line;

            while ((line = br.readLine()) != null) {
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
     * write content in a file
     *
     * @param contents  contents to write
     * @param outputDir output directory path
     * @param fileName  file name
     * @return isSuccess
     */
    public static boolean writeFile(Collection<String> contents, String outputDir, String fileName) {
        if (contents == null || contents.isEmpty()) {
            System.err.println("FileUtil : Writing a file is failed since contents is null");
            return false;
        }

        if (outputDir == null || fileName == null) {
            System.err.println("FileUtil : Writing a file is failed since outputDir or fileName is null");
            return false;
        }

        File resDir = new File(outputDir);

        if (!resDir.exists()) {
            if (!resDir.mkdir()) {
                System.err.println("FileUtil : Making a directory is failed.");
            }
        } else {
            if (!resDir.isDirectory()) {
                System.err.println("FileUtil : outputDir is not a directory path - " + outputDir);
            }
        }

        File writingFile = new File(resDir, fileName);

        int redundant = 0;

        while (writingFile.exists()) {
            String avoidRedundantFileName = fileName.substring(0, fileName.lastIndexOf("."));
            String originExtension = fileName.substring(fileName.lastIndexOf("."));
            writingFile = new File(resDir, avoidRedundantFileName + "(" + redundant + ")" + originExtension);
            redundant++;
        }

        try {
            if (writingFile.createNewFile()) {
                BufferedWriter bw = new BufferedWriter(new FileWriter(writingFile));
                for (String content : contents) {
                    bw.write(content);
                    bw.newLine();
                }
                bw.close();
                return true;
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }

    /**
     * Append contents in file
     *
     * @param contents contents to append
     * @param filePath file path
     * @return isSuccess
     */
    public static boolean appendContents(Collection<String> contents, String filePath) {
        if (contents == null || contents.isEmpty()) {
            System.out.println("FileUtil : Cannot append contents since contents are empty or null");
        }

        if (filePath == null) {
            System.err.println("FileUtil : Appending error - filePath is null");
        }

        File file = new File(filePath);

        if (!file.exists()) {
            System.err.println("FileUtil : Appending error - this file dose not exist > " + filePath);
        }

        try {
            BufferedWriter bw = new BufferedWriter(new FileWriter(file, true));

            for (String content : contents) {
                bw.write(content);
                bw.newLine();
            }

            bw.close();
            return true;
        } catch (IOException e) {
            e.printStackTrace();
        }
        return false;
    }
}