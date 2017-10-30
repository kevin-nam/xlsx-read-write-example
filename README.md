# xlsx-read-write-example
By: Kevin Nam (kevin.nam@mail.mcgill.ca)
For: Robert Sabourin

## Description
Example project for reading and writing into a premade excel file. The excel file being used can be found under src/main/resources. To be specific, the excel file has two sheets/tab: the first having no use, and the second having a (value, rate, converted value) as columns. The third column (converted value) has a formula associated with it such that (C2 = A2 * B2).

  In this application demo, we will be going into the second tab and grabbing the value in the position A2 (the value). Then we will be modifying the value in B2 (the rate) and saving the excel file.
  
  **The excel file can be found under src/main/java/resources/excel-file.xlsx**

## Requirements
1. JDK v1.8 [Link](http://www.oracle.com/technetwork/java/javase/downloads/jdk8-downloads-2133151.html)
2. Maven v3.3+ [Link](https://maven.apache.org/download.cgi)
3. IntelliJ 2017.1.3 Community Edition

## Dependencies Used

Using Maven, all dependencies can be easily acquired and used. Refer to */pom.xml* for more details.

```
<dependencies>
    <dependency>
          <groupId>org.apache.poi</groupId>
          <artifactId>poi-ooxml</artifactId>
          <version>3.17</version>
        </dependency>
        <dependency>
          <groupId>junit</groupId>
          <artifactId>junit</artifactId>
          <version>4.12</version>
        </dependency>
</dependencies>
```

## Classes Explained

- com.robsab.xlsx.**XlsxReaderWriter**
  * Class that facilitates reading and writing into an excel-file
- com.robsab.**Application**
  * Main class to demo the reading and writing into an excel-file. Run the main method of this class to demo.

## Tests Explained

One test (found under src/test/robsab/xlsx) have been implemented to express the use of reading and writing into an excel file.

1. **testReadAndWrite**
* The following test does as follows:
  1) Gets the rate value at B2
  2) Sets a new rate value at B2
  3) Gets the new rate value at B2
  4) Compares and asserts that the new rate value is properly written
  5) Gets the converted value at C2; a formula cell (such that C2 = A2 * B2)
  6) Calculates an expected converted value
  7) Compares and asserts the converted value at C2 and the expected converted value are equal

## References

- [Apache POI for HSSF and XSSF Quick Guide](https://poi.apache.org/spreadsheet/quick-guide.html)

