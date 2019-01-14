package com.zaptrapp.excelreadwritetest;

import android.content.Context;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.widget.Toast;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class MainActivity extends AppCompatActivity {
    public static final String TAG = MainActivity.class.getSimpleName();
    ArrayList<String> listProductName;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);

        addRandomArrayList();
//        if(saveExcelFile(this,"excel5.xls")){
//            Toast.makeText(this, "Saved", Toast.LENGTH_SHORT).show();
//        }

        /*Write to random area*/
        Sheet sheet = ExcelUtils.initialiseSheet(this,"Sales");
        if(ExcelUtils.writeValueInLocation(sheet,8,37,"Poornima")){
            Log.d(TAG, "onCreate: Poornima successful");
        }
        ExcelUtils.writeValueInLocation(sheet,90,5,"Nishanth");

//        int searchLocation = ExcelUtils.getRowLocationForValue("Nishanth",this,"excel5.xls","Sales",0,5,100);


        int rowNumber =90;
        int columnNumber=5;
        String value="";


        ExcelUtils.findAndRemove("Nishanth",this,"excel5.xls",sheet,"Sales",0,5,100);


        for(int i=0;i<200;i++){

            ExcelUtils.writeValueInLocation(sheet,i,7,"String "+i );
        }

        ExcelUtils.findAndRemove("String 67",this,"excel5.xls",sheet,"Sales",0,7,100);

        int searchLocation = ExcelUtils.findNearestEmpty(this,"excel5.xls","Sales",0,7,100);
        Toast.makeText(this, "Location is "+searchLocation, Toast.LENGTH_SHORT).show();

        Log.d(TAG, "onCreate: searchLocation is "+searchLocation);

        ExcelUtils.writeSheetToFile(this,"excel5.xls",sheet.getWorkbook());

//        Log.d(TAG, "onCreate: Location "+searchLocation);


    }




    private void addRandomArrayList() {
        listProductName = new ArrayList<>();
        for(int i=0;i<100;i++){
            listProductName.add(i,"String "+i);
        }
    }


    @Override
    protected void onPause() {
        super.onPause();
        listProductName.clear();
    }


    private boolean saveExcelFile(Context context, String fileName) {

        boolean success = false;

        try {
            Sheet sheet1 = ExcelUtils.initialiseSheet(context,"Sales");


            ExcelUtils.writeValueInLocation(sheet1,0,0, String.valueOf(34));

            Row row = sheet1.createRow(Integer.parseInt(ExcelUtils.getValueFromLocation(context,fileName,"Sales",0,0)));
//            row = sheet1.createRow(1);
            for(int i=0;i<listProductName.size();i++){
                ExcelUtils.writeValueInLocation(sheet1,0,i,listProductName.get(i));
                Log.d(TAG, "forLoop: "+listProductName.get(i));
            }


            ExcelUtils.writeSheetToFile(context, fileName, sheet1.getWorkbook());


        } catch (Exception e) {
            e.printStackTrace();
            Toast.makeText(context, "Error in saveExcelFile", Toast.LENGTH_SHORT).show();
            Log.d(TAG, "saveExcelFile: "+e.toString());
        }
        return success;
    }


}
