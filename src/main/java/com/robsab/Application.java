package com.robsab;

import com.robsab.xlsx.XlsxReaderWriter;

/**
 * Main application class
 *
 * Run the main method to demo reading and writing to an excel file.
 *
 * The excel file being used can be found under src/main/resources.
 * To be specific, the excel file has two sheets/tab: the first having no use, and the second having a (value, rate, converted value) as columns.
 * The third column (converted value) has a formula associated with it such that (C2 = A2 * B2).
 *
 * In this application demo, we will be going into the second tab and grabbing the value in the position A2 (the value).
 * Then we will be modifying the value in B2 (the rate) and saving the excel file.
 */
public class Application {

  private final static String PATH_TO_FILE = "src/main/resources/excel-file.xlsx";

  public static void main(String[] args) {

    // Instantiate the excel reader-writer to the file path
    XlsxReaderWriter xlsxReaderWriter = new XlsxReaderWriter(PATH_TO_FILE);

    // Get the value at sheet 1, row 1, column 0
    // In other words, the second tab/sheet at position A2
    String value = xlsxReaderWriter.getValueAtSheetRowAndColumn(1, 1, 0);
    System.out.println("Value at A2 in the second tab/sheet = " + value);

    // Write a value at sheet 1, row 1, column 1
    // In other words, writes a value in the second tab/sheet at position B2
    xlsxReaderWriter.writeValueAtSheetRowAndColumn(1, 1, 1, "0.77");
  }

}
