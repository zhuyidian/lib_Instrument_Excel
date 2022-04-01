package com.dunn.instrument.excel.sample;

import android.app.Activity;
import android.content.Intent;
import android.os.Bundle;
import android.util.Log;
import android.view.View;

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
        ApiExcel.excelSubmit();
    }

    @Override
    protected void onDestroy() {
        super.onDestroy();
    }
}
