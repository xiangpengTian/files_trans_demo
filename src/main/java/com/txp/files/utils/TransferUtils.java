package com.txp.files.utils;

public class TransferUtils {

    //判断当前系统是windows还是Linnux,是win 返回false

    public static boolean checkIsLinux(){

        String os = System.getProperty("os.name");
        if(os.toLowerCase().startsWith("win")){
            System.out.println("当前系统为Windows");
            return false;

        }
        return true;
    }
}
