package com.robsab.xlsx;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * The Excel Reader and Writer class
 * Has two primary methods for reading and writing an excel file:
 *      1) Writing a value at a specific sheet/tab, row and column
 *      2) Reading a value at a specific sheet/tab, row and column
 */
public class XlsxReaderWriter {

  private FileInputStream fileInputStream;
  private String pathToFile;

  public XlsxReaderWriter(String pathToFile) {
    this.pathToFile = pathToFile;
  }

  /** Method to write a value at a specific sheet/tab, row and column.
  * Important to note that the first tab has a sheet number index of 0.
  * Similarly the first row and first column both have indexes of 0
  * For example, position A1 would have row 0 and column 0
  */
  public void writeValueAtSheetRowAndColumn(int sheetNum, int rowNum, int colNum, String value) {

    try {
      System.out.println("Writing value " + value + " at sheet " + sheetNum + ", row " + rowNum + " & col " + colNum);

      // Open Excel File for reading
      openReadForExcelFile();

      // Open up the workbook
      XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

      // Get the tab/sheet
      XSSFSheet sheet = workbook.getSheetAt(sheetNum);

      // Go to row and column to get the cell
      XSSFRow row = sheet.getRow(rowNum);
      XSSFCell cell = row.getCell(colNum);

      // Close Excel file for reading
      closeReadForExcelFile();

      // Set the cell value
      cell.setCellValue(value);

      // Write into the Excel File
      writeExcelFile(workbook);

      System.out.println("Done writing!");
    } catch(IOException e) {
      e.printStackTrace();
    }

  }

  /** Method to read a value at a specific sheet/tab, row and column.
   * Important to note that the first tab has a sheet number index of 0.
   * Similarly the first row and first column both have indexes of 0
   * For example, position A1 would have row 0 and column 0
   */
  public String getValueAtSheetRowAndColumn(int sheetNum, int rowNum, int colNum) {
    try {
      System.out.println("Getting value at sheet " + sheetNum + ", row " + rowNum + " & col " + colNum);

      // Open Excel file for reading
      openReadForExcelFile();

      // Open up workbook
      XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);

      // Get the tab/sheet
      XSSFSheet sheet = workbook.getSheetAt(sheetNum);

      // Go to row and column to get the cell
      XSSFRow row = sheet.getRow(rowNum);
      XSSFCell cell = row.getCell(colNum);

      // Close Excel file for reading
      closeReadForExcelFile();

      // Return the value of the cell
      // If cell is null/blank, return empty string
      if (cell == null) {
        return "";
      }
      // If cell is numeric, return the string value of the numeric cell
      else if (cell.getCellTypeEnum().equals(CellType.NUMERIC)) {
        return String.valueOf(cell.getNumericCellValue());
      }
      // If cell is a formula, re-evaluate and calculate the value
      // We need to do this or else we will get the cached formula value
      else if (cell.getCellTypeEnum().equals(CellType.FORMULA)) {
        return String.valueOf(workbook.getCreationHelper().createFormulaEvaluator().evaluate(cell).getNumberValue());
      }
      // Otherwise, just return the string cell value
      else {
        return cell.getStringCellValue();
      }
    } catch(IOException e) {
      e.printStackTrace();
    }

    // Return empty string if error occurs
    return "";
  }

  // Method to open the read channel for the excel file
  private void openReadForExcelFile() throws IOException {
    fileInputStream = new FileInputStream(new File(pathToFile));
  }

  // Method to close the read channel for the excel file
  private void closeReadForExcelFile() throws IOException {
    fileInputStream.close();
  }

  // Method to write into the excel file by workbook
  private void writeExcelFile(XSSFWorkbook workbook) throws IOException {
    FileOutputStream fileOutputStream = new FileOutputStream(new File(pathToFile));
    workbook.write(fileOutputStream);
    fileOutputStream.close();
  }

}
