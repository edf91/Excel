package org.wxd.excel.export;

import com.google.common.collect.Lists;
import com.google.common.io.Closer;
import com.google.common.io.Files;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;
import org.wxd.excel.handler.export.DefaultExportHandler;
import org.wxd.excel.model.Clazz;
import org.wxd.excel.model.ExcelDTO;
import org.wxd.excel.model.Student;
import org.wxd.excel.model.Teacher;
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
    private static Integer entityNum = 100;
    private static List<String> sheetTitles = Lists.newArrayList("stu","tea");

    @Test
    public void exportTest(){
        excelDTO = buildEntity();
        String contextPath = System.getProperty("user.dir");
        File templateFile = new File(contextPath + File.separator + excelFileName);
        File tempFile = new File(contextPath + File.separator + "temp." + Files.getFileExtension(templateFile.getAbsolutePath()));
        File resultFile = new File(contextPath + File.separator + exportResultFileName);
        try{
            Files.copy(templateFile, tempFile);
            Closer closer = Closer.create();
            BufferedOutputStream out = closer.register(new BufferedOutputStream(new FileOutputStream(resultFile)));
            Workbook workbook = ExcelUtil.buildWorkbookFromEntityOfFile(
                    excelDTO,
                    sheetTitles,
                    tempFile,
                    new DefaultExportHandler()
            );
            workbook.write(out);
            workbook.close();
            closer.close();
        }catch (Exception e){
            e.printStackTrace();
        }

    }


    public ExcelDTO buildEntity(){
        Clazz clazz = new Clazz();
        for(int i = 0; i < entityNum; i++){
            clazz.getStudents().add(new Student(i + 1,"wang" + i,"man" + i));
            clazz.getTeachers().add(new Teacher(i + 1,"xiao" + i,"women" + i));
        }
        excelDTO.setClazz(clazz);
        return excelDTO;
    }
}
