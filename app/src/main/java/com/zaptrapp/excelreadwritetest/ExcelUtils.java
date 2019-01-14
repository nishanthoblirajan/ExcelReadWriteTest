package com.zaptrapp.excelreadwritetest;

import android.content.Context;
import android.util.Log;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;


/*Zaptr Tech
 * Documentation
 * 1. Create new Sheet Variable and use ExcelUtils.initialisesheet(context, sheetname)
 * Remember the sheet name is not the .xls file but the sheet name in that file
 * 2. Use the ExcelUtils.setValueInLocation(sheetName, rowNumber, columnNumber and the value (in string)
 * 3. To write the data to the file use the ExcelUtils.writeSheetToFile(context, fileName, workbook)
 * (Normally use the sheet.getWorkbook() to get the corresponding workbook.
 * 4. To read from a cell use the getValueFromLocation(context, fileName, sheetName, rowNumber, columnNumber)
 * */
public class ExcelUtils {


    public static final String TAG = "ExcelUtils";

    public static Sheet initialiseSheet(Context context, String sheetName) {
        Workbook wb = new HSSFWorkbook();
        Cell c = null;
        Sheet sheet1 = null;
        try {
            sheet1 = wb.createSheet(sheetName);
            return sheet1;
        } catch (Exception e) {
            Toast.makeText(context, "Unable to create SheetName " + sheetName, Toast.LENGTH_SHORT).show();
            return null;
        }
    }


    public static boolean setValueInLocation(Sheet sheet, int rowNumber, int columnNumber, String value) {
        Row row = null;
        if (sheet.getRow(rowNumber) == null) {
            row = sheet.createRow(rowNumber);
        } else {
            row = sheet.getRow(rowNumber);
        }
        Cell column;
        if (row.getCell(columnNumber) == null) {
            column = row.createCell(columnNumber);
        } else {
            column = row.getCell(columnNumber);
            row.removeCell(column);
            column = row.createCell(columnNumber);
        }
        column.setCellValue(value);
        if (column.getStringCellValue().equals(value)) {
            Log.d(TAG, "setValueInLocation: " + value + " saved in cell " + rowNumber + ", " + columnNumber);
            return true;
        } else {
            return false;
        }
    }


    public static Workbook getSheets(Context context, String fileName) throws IOException {
        File file = new File(context.getExternalFilesDir(null), fileName);
        FileInputStream os = null;
        try {
            os = new FileInputStream(file);
            Log.w("FileUtils", "Writing file" + file);
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }
        return WorkbookFactory.create(file);
    }

    public static Sheet getSheetfromString(Context context, String fileName, String sheetName) {
        Sheet returnSheet;
        try {
            returnSheet = getSheets(context, fileName).getSheet(sheetName);
        } catch (IOException e) {
            returnSheet = null;
            Toast.makeText(context, "No sheet found", Toast.LENGTH_SHORT).show();
            e.printStackTrace();
        }
        return returnSheet;
    }


    public static String getValueFromLocation(Context context, String fileName, String sheetName, int rowNumber, int columnNumber) {
        String string = "";
        try {
            Workbook workbook = getSheets(context, fileName);
            Log.d(TAG, "readExcelFile: Workbook has " + workbook.getNumberOfSheets() + " Sheets");
            Log.d(TAG, "readExcelFile: Sheets are as follows");
            for (Sheet sheet : workbook) {
                Log.d(TAG, "readExcelFile: " + sheet.getSheetName());
            }
            Sheet sheet = workbook.getSheet(sheetName);
            if (sheet != null) {
                // Create a DataFormatter to format and get each cell's value as String
                DataFormatter dataFormatter = new DataFormatter();
                if (sheet.getRow(rowNumber) != null) {
                    Row row = sheet.getRow(rowNumber);
                    if (row.getCell(columnNumber) != null) {
                        Cell cell = row.getCell(columnNumber);
                        Log.d(TAG, "readExcelFile: Cell at " + rowNumber + ", " + columnNumber + " is " + dataFormatter.formatCellValue(cell));

                        string = dataFormatter.formatCellValue(cell);
                    } else {
                        Log.d(TAG, "getValueFromLocation: No Column Found");
                        string = "";
                    }
                } else {
                    Log.d(TAG, "getValueFromLocation: No Row Found");
                    string = "";
                }

            } else {
                Log.d(TAG, "getValueFromLocation: Sheet not found");
                Toast.makeText(context, "Sheet not found", Toast.LENGTH_SHORT).show();
            }
        } catch (IOException e) {
            e.printStackTrace();
            Log.d(TAG, "getValueFromLocation: " + e.getMessage());
        }

        return string;
    }


    public static int getRowLocationForValue(String searchString, Context context, String fileName, String sheetName, int startingRowNumber, int startingColumnNumber, int count) {
        int loopCount = startingRowNumber + count;
        for (int i = startingRowNumber; i < loopCount; i++) {
            String stringInThatCell = getValueFromLocation(context, fileName, sheetName, i, startingColumnNumber);
            if (stringInThatCell.equals(searchString)) {
                Log.d(TAG, "getRowLocationForValue: " + searchString + " found at " + i);
                return i;
            }
        }
        return 0;
    }

    public static void writeSheetToFile(Context context, String fileName, Workbook wb) {
        File file = new File(context.getExternalFilesDir(null), fileName);
        FileOutputStream os = null;

        try {
            os = new FileOutputStream(file);
            wb.write(os);
            Log.w("FileUtils", "Writing file" + file);
        } catch (IOException e) {
            Log.w("FileUtils", "Error writing " + file, e);
        } catch (Exception e) {
            Log.w("FileUtils", "Failed to save file", e);
        } finally {
            try {
                if (null != os)
                    os.close();
            } catch (Exception ex) {
            }
        }
    }

    public static int findNearestEmpty(Context context, String fileName, String sheetName, int startingRowNumber, int startingColumnNumber, int count) {
        int nearestEmpty = getRowLocationForValue("", context, fileName, sheetName, startingRowNumber, startingColumnNumber, count);
        Log.d(TAG, "findNearestEmpty: The nearest empty is " + nearestEmpty);
        return nearestEmpty;
    }

    public static void deleteValueFromLocation(Sheet sheet, int rowNumber, int columnNumber) {
        ExcelUtils.setValueInLocation(sheet, rowNumber, columnNumber, "");
    }

    public static void replaceValueInLocation(Sheet sheet, int rowNumber, int columnNumber, String newString) {
        ExcelUtils.setValueInLocation(sheet, rowNumber, columnNumber, newString);
    }

    public static boolean findAndRemove(String searchString, Context context, String fileName, Sheet sheet, String sheetName, int startingRowNumber, int startingColumnNumber, int count) {
        int location = getRowLocationForValue(searchString, context, fileName, sheetName, startingRowNumber, startingColumnNumber, count);
        if (location != 0) {
            deleteValueFromLocation(sheet, location, startingColumnNumber);
            Log.d(TAG, "findAndRemove: Deleted " + location + ", " + startingColumnNumber);
            Toast.makeText(context, "Deleted", Toast.LENGTH_SHORT).show();
            return true;
        } else {
            Log.d(TAG, "findAndRemove: Unable to find");
            Toast.makeText(context, "Unable to delete", Toast.LENGTH_SHORT).show();
            return false;
        }

    }

    public static boolean isLocationEmpty(Context context, String fileName, String sheetName, int rowNumber, int columnNumber) {
        if (getValueFromLocation(context, fileName, sheetName, rowNumber, columnNumber).equals("")) {
            return true;
        } else {
            return false;
        }
    }

    public static boolean addValueToNearestEmpty(Context context, String fileName, Sheet sheet, String sheetName, int startingRowNumber, int startingColumnNumber, int count, String value) {
        int emptyLocation = findNearestEmpty(context, fileName, sheetName, startingRowNumber, startingColumnNumber, count);
        if(isLocationEmpty(context, fileName, sheetName, emptyLocation, startingColumnNumber)) {
            if (emptyLocation != 0) {
                if (setValueInLocation(sheet, emptyLocation, startingColumnNumber, value)) {
                    return true;
                } else {
                    return false;
                }
            } else {
                Log.d(TAG, "addValueToNearestEmpty: Error");
                return false;

            }
        }else{
            return false;
        }

    }
}
