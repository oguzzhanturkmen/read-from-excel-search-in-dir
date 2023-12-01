package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class FolderSearch {
static  Set<String> set = new TreeSet<>();

    static class SearchPair {

        // Class to store the search pair, you may add more fields if you want to search more words in the directory name
        String firstString;
        String fourthString;

        SearchPair(String firstString, String fourthString) {
            this.firstString = firstString;
            this.fourthString = fourthString;
        }

        @Override
        public String toString() {
            return "SearchPair{" +
                    "firstString='" + firstString + '\'' +
                    ", fourthString='" + fourthString + '\'' +
                    '}';
        }
    }

    public static void main(String[] args) {
        try {
            // Path to the Excel file
            String excelFilePath = "<Please give your path for excel file >";
            List<SearchPair> searchPairs = readExcelFile(excelFilePath); // Reading the Excel file and getting the list of SearchPair objects

            // Directory to start searching from
            String directoryPath = "<Please give your path for parent directory to searcy>";



            for (SearchPair pair : searchPairs) { // Iterating all SearchPair objects
                searchDirectory(new File(directoryPath), pair ); // Calling the searchDirectory method
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        for (   String s : set) { // Iterating the set
            System.out.println(s); // Printing the result
        }
              {

        }
    }

    private static List<SearchPair> readExcelFile(String filePath) throws Exception {

        List<SearchPair> searchPairs = new ArrayList<>(); // List to store SearchPair objects

        FileInputStream file = new FileInputStream(new File(filePath)); // Creating a file input stream to read Excel workbook

        System.out.println(new File(filePath).exists()); // Checking if the file exists

        Workbook workbook = new XSSFWorkbook(file); // Creating Workbook instance that refers to .xlsx file

        Sheet sheet = workbook.getSheetAt(0); // Creating a Sheet object to retrieve the object

        for (Row row : sheet) { // Iterating all rows in the sheet
            if (row.getRowNum() >= 1) { // Skipping the first row (header)
                Cell firstColumnCell = row.getCell(0); // 1st column
                Cell fourthColumnCell = row.getCell(3); // 4th column
                if (firstColumnCell != null && fourthColumnCell != null) { // Checking for null values
                    searchPairs.add(new SearchPair(String.valueOf((int)firstColumnCell.getNumericCellValue()), fourthColumnCell.getStringCellValue().substring(0,5)));
                }
            }
        }
        workbook.close();
        return searchPairs; // Returning the list of SearchPair objects
    }

    private static void searchDirectory(File directory, SearchPair pair ) {
        File[] files = directory.listFiles(); // Getting all files and directories in the specified directory

        if (files != null) { // Checking if the directory is not empty
            for (File file : files) { // Iterating all files and directories

                if (file.isDirectory()) { // Checking if the current file is a directory
                    if (file.getName().contains(pair.firstString) && file.getName().contains(pair.fourthString)) { // Checking if the directory name contains the search string

                        String s = pair.firstString + " - " + pair.fourthString + " - " + file.getName(); // Creating a string to store the result
                        set.add(s); // Adding the result to the set
                        break;
                    }

                    // Search in subdirectories
                    searchDirectory(file, pair); // Calling the method recursively
                }
            }
        }
    }
}