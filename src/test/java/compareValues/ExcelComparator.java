package compareValues;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.*;

public class ExcelComparator {

    // Class to hold the column matches
    public static class ColumnMatch {
        private final int columnIndex1;
        private final int columnIndex2;

        public ColumnMatch(int columnIndex1, int columnIndex2) {
            this.columnIndex1 = columnIndex1;
            this.columnIndex2 = columnIndex2;
        }

        public int getColumnIndex1() {
            return columnIndex1;
        }

        public int getColumnIndex2() {
            return columnIndex2;
        }
    }

    // Helper method to read header mapping from the first row of the sheet
    private static Map<String, Integer> readHeaderMapping(Sheet sheet) {
        Row headerRow = sheet.getRow(0);
        Map<String, Integer> headerMap = new HashMap<>();

        for (Cell cell : headerRow) {
            String header = cell.getStringCellValue().trim();
            int columnIndex = cell.getColumnIndex();
            headerMap.put(header, columnIndex);
        }

        return headerMap;
    }

    // Helper method to read column mapping from the mapping sheet
    private static Map<String, String> readColumnMapping(Sheet mappingSheet) {
        Map<String, String> columnMapping = new HashMap<>();

        for (Row row : mappingSheet) {
            String sourceColumn = row.getCell(0).getStringCellValue().trim();
            String targetColumn = row.getCell(1).getStringCellValue().trim();
            columnMapping.put(sourceColumn, targetColumn);
        }

        return columnMapping;
    }

    // Helper method to match columns based on the explicit column mapping
    private static List<ColumnMatch> matchColumns(Map<String, Integer> headerMap1, Map<String, Integer> headerMap2,
                                                  Map<String, String> columnMapping) {
        List<ColumnMatch> columnMatches = new ArrayList<>();

        for (Map.Entry<String, String> entry : columnMapping.entrySet()) {
            String sourceColumn = entry.getKey();
            String targetColumn = entry.getValue();

            if (headerMap1.containsKey(sourceColumn) && headerMap2.containsKey(targetColumn)) {
                int columnIndex1 = headerMap1.get(sourceColumn);
                int columnIndex2 = headerMap2.get(targetColumn);
                columnMatches.add(new ColumnMatch(columnIndex1, columnIndex2));
            } else {
                System.out.println("Column mapping not found for '" + sourceColumn + "' or '" + targetColumn + "'.");
                columnMatches.add(null); // Add a placeholder to indicate the mismatch
            }
        }

        return columnMatches;
    }

    // Helper method to get the column header from the index
    private static String getColumnKey(Map<String, Integer> headerMap, int columnIndex) {
        for (Map.Entry<String, Integer> entry : headerMap.entrySet()) {
            if (entry.getValue() == columnIndex) {
                return entry.getKey();
            }
        }
        return null;
    }

    // Helper method to compare cell values
    private static boolean cellValuesMatch(Cell cell1, Cell cell2) {
        if (cell1 == null && cell2 == null) {
            return true;
        }
        if (cell1 == null || cell2 == null) {
            return false;
        }

        CellType cellType1 = cell1.getCellType();
        CellType cellType2 = cell2.getCellType();

        if (cellType1 != cellType2) {
            return false;
        }

        if (cellType1 == CellType.STRING) {
            return cell1.getStringCellValue().equals(cell2.getStringCellValue());
        } else if (cellType1 == CellType.NUMERIC) {
            return cell1.getNumericCellValue() == cell2.getNumericCellValue();
        } else if (cellType1 == CellType.BOOLEAN) {
            return cell1.getBooleanCellValue() == cell2.getBooleanCellValue();
        } else if (cellType1 == CellType.BLANK) {
            return true;
        }

        return false;
    }

    // Helper method to get cell value as string
    private static String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        CellType cellType = cell.getCellType();
        if (cellType == CellType.STRING) {
            return cell.getStringCellValue();
        } else if (cellType == CellType.NUMERIC) {
            return String.valueOf(cell.getNumericCellValue());
        } else if (cellType == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else if (cellType == CellType.BLANK) {
            return "";
        }

        return "";
    }
    private static void createHighlightedExcelFile(Sheet sheet1, Sheet sheet2, Map<String, Integer> headerMap1,
            Map<String, Integer> headerMap2, List<ColumnMatch> columnMatches,
            String outputPath) throws IOException {
Workbook workbookOutput = new XSSFWorkbook();
Sheet sheetOutput = workbookOutput.createSheet("Mismatches");

// CellStyle for highlighting mismatched cells
CellStyle styleMismatch = workbookOutput.createCellStyle();
styleMismatch.setFillForegroundColor(IndexedColors.RED.getIndex());
styleMismatch.setFillPattern(FillPatternType.SOLID_FOREGROUND);

int rowIndexOutput = 0;
int tagId = 1;

Row headerRowOutput = sheetOutput.createRow(rowIndexOutput);
headerRowOutput.createCell(0).setCellValue("Tag ID");
headerRowOutput.createCell(1).setCellValue("Field Name");
headerRowOutput.createCell(2).setCellValue("Value in File 1");
headerRowOutput.createCell(3).setCellValue("Value in File 2");

rowIndexOutput++;

for (ColumnMatch columnMatch : columnMatches) {
if (columnMatch == null) {
continue; // Skip the placeholder for mismatched mapping
}

int columnIndex1 = columnMatch.getColumnIndex1();
int columnIndex2 = columnMatch.getColumnIndex2();

String columnKey1 = getColumnKey(headerMap1, columnIndex1);
String columnKey2 = getColumnKey(headerMap2, columnIndex2);

Iterator<Row> rowIterator1 = sheet1.iterator();
Iterator<Row> rowIterator2 = sheet2.iterator();

// Skip header row
rowIterator1.next();
rowIterator2.next();

while (rowIterator1.hasNext() && rowIterator2.hasNext()) {
Row row1 = rowIterator1.next();
Row row2 = rowIterator2.next();

Cell cell1 = row1.getCell(columnIndex1);
Cell cell2 = row2.getCell(columnIndex2);

String cellValue1 = getCellValueAsString(cell1);
String cellValue2 = getCellValueAsString(cell2);

if (!cellValuesMatch(cell1, cell2)) {
Row mismatchRowOutput = sheetOutput.createRow(rowIndexOutput);
mismatchRowOutput.createCell(0).setCellValue(tagId);
mismatchRowOutput.createCell(1).setCellValue(columnKey1);
mismatchRowOutput.createCell(2).setCellValue(cellValue1);
mismatchRowOutput.createCell(3).setCellValue(cellValue2);

Cell mismatchCell1 = mismatchRowOutput.getCell(1);
Cell mismatchCell2 = mismatchRowOutput.getCell(2);
Cell mismatchCell3 = mismatchRowOutput.getCell(3);

mismatchCell1.setCellStyle(styleMismatch);
mismatchCell2.setCellStyle(styleMismatch);
mismatchCell3.setCellStyle(styleMismatch);

rowIndexOutput++;
}
}

tagId++;
}

// Auto-size columns for better visibility of the contents
for (int i = 0; i < headerRowOutput.getLastCellNum(); i++) {
sheetOutput.autoSizeColumn(i);
}

// Write the output workbook to a file
try (FileOutputStream fos = new FileOutputStream(outputPath)) {
workbookOutput.write(fos);
}

workbookOutput.close();
}




 


    // Helper method to compare the two Excel sheets using all the column matches
    public static boolean compareSheets(Sheet sheet1, Sheet sheet2, Map<String, Integer> headerMap1,
                                         Map<String, Integer> headerMap2, List<ColumnMatch> columnMatches) {
        boolean hasMismatch = false;

        for (int i = 0; i < columnMatches.size(); i++) {
            ColumnMatch columnMatch = columnMatches.get(i);

            if (columnMatch == null) {
                hasMismatch = true;
                continue;
            }

            int columnIndex1 = columnMatch.getColumnIndex1();
            int columnIndex2 = columnMatch.getColumnIndex2();

            String columnKey1 = getColumnKey(headerMap1, columnIndex1);
            String columnKey2 = getColumnKey(headerMap2, columnIndex2);

            Iterator<Row> rowIterator1 = sheet1.iterator();
            Iterator<Row> rowIterator2 = sheet2.iterator();

            // Skip header row
            rowIterator1.next();
            rowIterator2.next();

            while (rowIterator1.hasNext() && rowIterator2.hasNext()) {
                Row row1 = rowIterator1.next();
                Row row2 = rowIterator2.next();

                Cell cell1 = row1.getCell(columnIndex1);
                Cell cell2 = row2.getCell(columnIndex2);

                if (!cellValuesMatch(cell1, cell2)) {
                    hasMismatch = true;
                    System.out.println("Comparing: " + columnKey1 + " and " + columnKey2);
                    System.out.println("Cell 1: " + getCellValueAsString(cell1));
                    System.out.println("Cell 2: " + getCellValueAsString(cell2));
                    System.out.println();
                }
            }
        }

        return !hasMismatch;
    }

    
    
    
    
    // Main method
    public static void main(String[] args) {
    	 String excelFile1Path = "/Users/birendra/Desktop/file.xlsx";
         String excelFile2Path = "/Users/birendra/Desktop/target.xlsx";
         String mappingFilepath = "/Users/birendra/Desktop/mapping.xlsx";

        try (FileInputStream fis1 = new FileInputStream(excelFile1Path);
             FileInputStream fis2 = new FileInputStream(excelFile2Path);
             FileInputStream fisMapping = new FileInputStream(mappingFilepath)) {

            Workbook workbook1 = new XSSFWorkbook(fis1);
            Workbook workbook2 = new XSSFWorkbook(fis2);
            Workbook workbookMapping = new XSSFWorkbook(fisMapping);

            // Get the first sheet from each Excel file
            Sheet sheet1 = workbook1.getSheetAt(0);
            Sheet sheet2 = workbook2.getSheetAt(0);
            Sheet mappingSheet = workbookMapping.getSheetAt(0);

            // Read header mapping from the first row of the sheets
            Map<String, Integer> headerMap1 = readHeaderMapping(sheet1);
            Map<String, Integer> headerMap2 = readHeaderMapping(sheet2);

            // Read column mapping from the mapping sheet
            Map<String, String> columnMapping = readColumnMapping(mappingSheet);

            // Match columns based on the explicit column mapping
            List<ColumnMatch> columnMatches = matchColumns(headerMap1, headerMap2, columnMapping);

            // Perform comparison between the two Excel sheets using the best column matches
            if (!compareSheets(sheet1, sheet2, headerMap1, headerMap2, columnMatches)) {
                System.out.println("Excel sheets have different values.");

                // Create a new Excel file to highlight the mismatches
                try {
                    String outputPath = "/Users/birendra/Desktop/mismatches.xlsx"; // Update with the desired output path
                    createHighlightedExcelFile(sheet1, sheet2, headerMap1, headerMap2, columnMatches, outputPath);
                    System.out.println("Mismatched values have been highlighted in the new Excel file: " + outputPath);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            } else {
                System.out.println("Excel sheets are identical.");
            }

            // Close the workbooks
            workbook1.close();
            workbook2.close();
            workbookMapping.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
