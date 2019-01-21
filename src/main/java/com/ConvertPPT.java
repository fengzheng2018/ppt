package com;

import org.apache.poi.hslf.usermodel.HSLFSlide;
import org.apache.poi.hslf.usermodel.HSLFSlideShow;

import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

/**
 * FileName: ConvertPPT
 * Create by fengzheng on 2019/1/21 21:22
 * Description: ppt格式转图片
 */
public class ConvertPPT {

    private File pptFile;
    private String targetPath;

    public ConvertPPT(File pptFile, String targetPath) {
        this.pptFile = pptFile;
        this.targetPath = targetPath;
    }

    public boolean doPPTToImage(){
        try{
            // 从流加载文件
            FileInputStream inputStream = new FileInputStream(this.pptFile);
            HSLFSlideShow ppt = new HSLFSlideShow(inputStream);
            inputStream.close();

            // 获取ppt页数
            Dimension pgSize = ppt.getPageSize();

            int idx = 1;
            for(HSLFSlide slide : ppt.getSlides()){
                BufferedImage image = new BufferedImage(pgSize.width,pgSize.height,BufferedImage.TYPE_INT_BGR);
                Graphics2D graphics = image.createGraphics();

                // 清空绘图区域
                graphics.setPaint(Color.white);
                graphics.fill(new Rectangle2D.Float(0,0,pgSize.width,pgSize.height));

                // 渲染
                slide.draw(graphics);

                // 检查文件是否存在
                File targetFile = new File(this.targetPath + File.separator + idx +".png");
                if(targetFile.exists()){
                    idx ++;
                    continue;
                }

                // 输出保存
                FileOutputStream outputStream = new FileOutputStream(targetFile);
                javax.imageio.ImageIO.write(image,"png",outputStream);

                outputStream.close();

                idx ++;
            }
        }catch (java.io.FileNotFoundException e){
            System.out.println("文件未找到！");
            return false;
        }catch (java.io.IOException e){
            System.out.println("加载文件异常，检查文件格式！");
            return false;
        }
        return true;
    }
}
