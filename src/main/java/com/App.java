package com;

import java.io.File;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        String targetPath = "E:/temp/pptText/";
        File file = new File("E:/temp/pptText/1.ppt");

        ConvertPPT convertPPT = new ConvertPPT(file,targetPath);
        boolean result = convertPPT.doPPTToImage();

        if(result){
            System.out.println("转换成功");
        }else{
            System.out.println("转换失败");
        }
    }
}
