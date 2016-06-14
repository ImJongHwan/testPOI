package org.poi.Util;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.nio.charset.StandardCharsets;
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
     * @return file absolute path
     */
    public static String writeFile(Collection<String> contents, String outputDir, String fileName) {
//        if (contents == null || contents.isEmpty()) {
//            System.err.println("FileUtil : Writing a file is failed since contents is null");
//            return null;
//        }

        if (outputDir == null || fileName == null) {
            System.err.println("FileUtil : Writing a file is failed since outputDir or fileName is null");
            return null;
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
                return writingFile.getAbsolutePath();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
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

    /**
     * read Resource file
     *
     * @param mainClass main class, ex) Main.class
     * @param resourceFilePath resource file path
     * @return read list
     * @throws IOException
     */
    public static List<String> readResourceFile(Class mainClass, String resourceFilePath) throws IOException{
        InputStream is = mainClass.getClassLoader().getResourceAsStream(resourceFilePath);

        if(is == null){
            System.err.println("FileUtil : Cannot read a resource file since path is wrong - " + resourceFilePath);
            return null;
        }

        BufferedReader br = new BufferedReader(new InputStreamReader(is, StandardCharsets.UTF_8));
        List<String> lines = new ArrayList<>();
        String line;

        while((line = br.readLine()) != null){
            lines.add(line);
        }

        br.close();

        return lines;
    }

    /**
     * Read excel file
     * if filePath is null, return is null.
     *
     * @param filePath Excel file path
     * @return XSSFWorkbook
     * @throws IOException
     */
    public static XSSFWorkbook readExcelFile(String filePath) throws IOException{
        if(filePath == null){
            System.err.println("FileUtil : Reading a excel file is failed since filePath is null");
            return null;
        }
        return new XSSFWorkbook(XSSFWorkbook.openPackage(filePath));
    }
}