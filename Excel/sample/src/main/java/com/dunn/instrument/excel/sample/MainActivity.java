package com.dunn.instrument.excel.sample;

import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;
import android.util.Log;
import android.view.View;

import com.dunn.instrument.excel.ExcelDeal;

import java.io.File;
import java.io.FileInputStream;

public class MainActivity extends Activity {
    public static String TAG = "MainActivity";

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
    }

    public void click(View view) {
        Log.i(TAG,"click:");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION1_COL,"1");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION2_COL,"2");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION3_COL,"3");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION4_COL,"4");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION5_COL,"5");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION6_COL,"6");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION7_COL,"7");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION8_COL,"8");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION9_COL,"9");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION10_COL,"10");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION11_COL,"11");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION12_COL,"12");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION13_COL,"13");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION14_COL,"14");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION15_COL,"15");
        ApiExcel.getInfo().setFunction1Value(ExcelDeal.FUNCTION16_COL,"16");
        ApiExcel.excelSubmit();
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
    }
}
