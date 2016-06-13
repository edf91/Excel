package org.wxd.excel.export;

import com.google.common.collect.Lists;
import com.google.common.io.Closer;
import com.google.common.io.Files;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.wxd.excel.handler.export.DefaultExportSyncHandler;
import org.wxd.excel.model.*;
import org.wxd.excel.utils.ExcelUtil;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.FileOutputStream;
import java.util.List;

/**
 * @Description:
 * @Copyright: Copyright 2012 ShenZhen DSE Corporation
 * @Company: 深圳市东深电子股份有限公司
 * @Author : wangxd
 * @Date: 2016-5-10
 * @Version 1.0
 */
public class ExcelExportTest {

    private static String excelFileName = "stu_tea_test.xlsx";
    private static String exportResultFileName = "stu_tea_test_result.xlsx";
    private static ExcelDTO excelDTO = new ExcelDTO();
    private static Integer entityNum = 1000;
    private static List<String> sheetTitles = Lists.newArrayList("stu","tea","stu2","stu3","stu4","stu5","stu6","stu7","stu8","stu9","stu10",
            "stu11","stu12","stu13","stu14","stu15");

    @SuppressWarnings("Duplicates")
    @Test
    public void poiExportTest(){
        excelDTO = buildEntity();
        String contextPath = System.getProperty("user.dir");
        File templateFile = new File(contextPath + File.separator + excelFileName);
        File resultFile = new File(contextPath + File.separator + exportResultFileName);
        try{
            Closer closer = Closer.create();
            BufferedOutputStream out = closer.register(new BufferedOutputStream(new FileOutputStream(resultFile)));
            Workbook workbook = ExcelUtil.buildWorkbookFromEntityOfFile(
                    excelDTO,
                    sheetTitles,
                    templateFile,
                    new DefaultExportSyncHandler()
            );
            workbook.write(out);
            closer.close();
        }catch (Exception e){
            e.printStackTrace();
        }

    }


//    @SuppressWarnings("Duplicates")
//    @Test
//    public void jxlExportTest(){
//        excelDTO = buildEntity();
//        String contextPath = System.getProperty("user.dir");
//        File templateFile = new File(contextPath + File.separator + excelFileName);
//        File tempFile = new File(contextPath + File.separator + "temp." + Files.getFileExtension(templateFile.getAbsolutePath()));
//        File resultFile = new File(contextPath + File.separator + exportResultFileName);
//        try{
//            Files.copy(templateFile, tempFile);
//            Closer closer = Closer.create();
//            BufferedOutputStream out = closer.register(new BufferedOutputStream(new FileOutputStream(resultFile)));
//            Workbook workbook = ExcelUtil.buildWorkbookFromEntityOfFile(
//                    excelDTO,
//                    sheetTitles,
//                    tempFile,
//                    new DefaultExportSyncHandler()
//            );
//            workbook.write(out);
//            workbook.close();
//            closer.close();
//        }catch (Exception e){
//            e.printStackTrace();
//        }
//
//    }


    public ExcelDTO buildEntity(){
        Clazz clazz = new Clazz();
        for(int i = 0; i < entityNum; i++){
            clazz.getStudents().add(new Student(i + 1,"wang" + i,"man" + i));
            clazz.getStudents2().add(new Student2(i + 1,"wang" + i,"man" + i));
            clazz.getStudents3().add(new Student3(i + 1,"wang" + i,"man" + i));
            clazz.getStudents4().add(new Student4(i + 1,"wang" + i,"man" + i));
            clazz.getStudents5().add(new Student5(i + 1,"wang" + i,"man" + i));
            clazz.getStudents6().add(new Student6(i + 1,"wang" + i,"man" + i));
            clazz.getStudents7().add(new Student7(i + 1,"wang" + i,"man" + i));
            clazz.getStudents8().add(new Student8(i + 1,"wang" + i,"man" + i));
            clazz.getStudents9().add(new Student9(i + 1,"wang" + i,"man" + i));
            clazz.getStudents10().add(new Student10(i + 1,"wang" + i,"man" + i));
            clazz.getStudents11().add(new Student11(i + 1,"wang" + i,"man" + i));
            clazz.getStudents12().add(new Student12(i + 1,"wang" + i,"man" + i));
            clazz.getStudents13().add(new Student13(i + 1,"wang" + i,"man" + i));
            clazz.getStudents14().add(new Student14(i + 1,"wang" + i,"man" + i));
            clazz.getStudents15().add(new Student15(i + 1,"wang" + i,"man" + i));
            clazz.getTeachers().add(new Teacher(i + 1,"xiao" + i,"women" + i));
        }
        excelDTO.setClazz(clazz);
        return excelDTO;
    }
}
