package com.dunn.instrument.excel.aop;

import android.util.Log;

import org.aspectj.lang.ProceedingJoinPoint;
import org.aspectj.lang.Signature;
import org.aspectj.lang.annotation.Around;
import org.aspectj.lang.annotation.Aspect;

/**
 * @ClassName: ExcelAspect
 * @Author: ZhuYiDian
 * @CreateDate: 2022/4/10 17:28
 * @Description:
 */
@Aspect
public class ExcelAspect {
    private static final String TAG = "ExcelAspect";

    /**
     * 监听com.coocaa.os.ccosservice.MainApp类中的所有方法
     * @param joinPoint
     */
    @Around("call(* com.coocaa.os.ccosservice.MainApp.**(..))")
    public void getTime(ProceedingJoinPoint joinPoint){
        Signature signature = joinPoint.getSignature();  //获取方法签名
        String name = signature.toShortString();  //获取方法所有信息
        long time = System.currentTimeMillis();
        try{
            joinPoint.proceed();
        }catch (Throwable throwable){
            throwable.printStackTrace();
        }
        Log.i(TAG,"getTime: name="+name+", cost="+(System.currentTimeMillis()-time)+"MS");
    }
}
