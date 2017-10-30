package com.robsab.xlsx;

import org.junit.Assert;
import org.junit.Test;

/**
 * Tests for XlsxReaderWriter
 *
 * The following test does as follows:
 *    1) Gets the rate value at B2
 *    2) Sets a new rate value at B2
 *    3) Gets the new rate value at B2
 *    4) Compares and asserts that the new rate value is properly written
 *    5) Gets the converted value at C2; a formula cell (such that C2 = A2 * B2)
 *    6) Calculates an expected converted value
 *    7) Compares and asserts the converted value at C2 and the expected converted value are equal
 */
public class XlsxReaderWriterTest {

  private final static String PATH_TO_FILE = "src/main/resources/excel-file.xlsx";

  @Test
  public void testReadAndWrite() {
    // Instantiate the excel reader-writer to the file path
    XlsxReaderWriter xlsxReaderWriter = new XlsxReaderWriter(PATH_TO_FILE);

    // Get the rate value at sheet 1, row 1, column 1
    // In other words, the second tab/sheet at position B2
    String rateValue = xlsxReaderWriter.getValueAtSheetRowAndColumn(1, 1, 1);
    System.out.println("Rate value at B2 in the second tab/sheet = " + rateValue);

    // Write a new rate value at sheet 1, row 1, column 1
    // In other words, writes a value in the second tab/sheet at position B2
    String newRateValue = "0.66";
    xlsxReaderWriter.writeValueAtSheetRowAndColumn(1, 1, 1, newRateValue);

    // Get the rate value again at sheet 1, row 1, column 1
    // In other words, the second tab/sheet at position B2
    String modifiedRateValue = xlsxReaderWriter.getValueAtSheetRowAndColumn(1, 1, 1);
    System.out.println("Rate value at B2 in the second tab/sheet after modification = " + modifiedRateValue);

    // Check that the rate value post-modified is equal to the new desired rate value
    Assert.assertEquals(modifiedRateValue, newRateValue);

    // Get the converted value at sheet 1, row 1, column 2
    // In other words, the second tab/sheet at position C2
    String convertedValue = xlsxReaderWriter.getValueAtSheetRowAndColumn(1, 1, 2);
    System.out.println("Converted value at C2 in the second tab/sheet after modification = " + convertedValue);

    // Get the value at sheet 1, row 1, column 0
    // In other words, the second tab/sheet at position A2
    // And calculate the expected converted value
    String value = xlsxReaderWriter.getValueAtSheetRowAndColumn(1, 1, 0);
    System.out.println("Value at A2 in the second tab/sheet = " + value);
    double expectedConvertedValue = Double.valueOf(value) * Double.valueOf(newRateValue);

    // Check that the converted value on the excel is equal to the expected calculated converted value
    Assert.assertEquals(String.valueOf(expectedConvertedValue), convertedValue);
  }

}
