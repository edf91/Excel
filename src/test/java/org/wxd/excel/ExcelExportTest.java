//package org.wxd.excel;
//
//
//import com.google.common.collect.Lists;
//import com.google.common.collect.Maps;
//import com.google.common.io.Closer;
//import org.apache.poi.ss.usermodel.Row;
//import org.apache.poi.ss.usermodel.Sheet;
//import org.apache.poi.ss.usermodel.Workbook;
//import org.apache.poi.ss.usermodel.WorkbookFactory;
//import org.wxd.excel.bean.ExcelTemplate;
//import org.wxd.excel.exception.ExcelException;
//
//import java.io.BufferedOutputStream;
//import java.io.File;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.math.BigDecimal;
//import java.util.List;
//import java.util.Map;
//
///**
// * @Description: 导出测试类
// * @Author : wangxd
// * @Date: 2016-3-2
// * @Version 1.0
// */
//public class ExcelExportTest {
//
//    public static ExcelExportTest test = new ExcelExportTest();
//    public static String path = "E:\\Develop\\Code\\IDEAProject\\ynfxkhbb\\WebRoot\\ynfxkhbb\\file\\reportTemp\\refactory\\FloodTemplate.xls";
//    public static String outPath = "E:\\Develop\\Code\\IDEAProject\\ynfxkhbb\\WebRoot\\ynfxkhbb\\file\\reportTemp\\refactory\\FloodTemplate2.xls";
//
//    public static List<FloodBaseInfoDTO> floodBaseInfos = Lists.newArrayList();
//    public static List<FloodNlmyyDTO> floodNlmyys = Lists.newArrayList();
//    public static ReportRequestDTO reportRequestDTO = new ReportRequestDTO();
//    public static FloodReportDTO floodReportDTO = new FloodReportDTO();
//
//    static {
//        for (int i = 0; i < 3; i++) {
//            FloodBaseInfoDTO baseInfo = new FloodBaseInfoDTO();
//            FloodNlmyyDTO floodNlmyy = new FloodNlmyyDTO();
//            floodNlmyy.setCorpCzCount(new BigDecimal(1 + i));
//            floodNlmyy.setCorpReduction(new BigDecimal(1 + i));
//            baseInfo.setDistrictName("测试" + i);
//            baseInfo.setCityNum(i);
//            baseInfo.setDevolveNum(i);
//            baseInfo.setRangeXianShiQu(i);
//            baseInfo.setRangeXianShiQu2(i + 1);
//            floodBaseInfos.add(baseInfo);
//            floodNlmyys.add(floodNlmyy);
//        }
//        floodReportDTO.setFloodBaseInfos(floodBaseInfos);
//        floodReportDTO.setFloodNlmyys(floodNlmyys);
//        reportRequestDTO.setFlood(floodReportDTO);
//    }
//
//    public File buildFileOfPath(String path) {
//        return new File(path);
//    }
//
//    public Workbook buildWorkbookOfTemplate(File template) {
//        try {
//            return WorkbookFactory.create(template);
//        } catch (Exception e) {
//            throw new ExcelException(e.getMessage(), e);
//        }
//    }
//
//    public Workbook getWorkbook() {
//        return test.buildWorkbookOfTemplate(test.buildFileOfPath(path));
//    }
//
//    public Workbook putToWorkbook() throws IOException {
//        Workbook workbook = this.getWorkbook();
//        List<ExcelTemplate> excelTemplates = null; //YNUtil.readReportRequestDTOAsTemplates(reportRequestDTO);
//        Sheet sheet = null;
//        Row row = null;
//        Map<String,Integer> hasDealIndexMap = Maps.newConcurrentMap();
//        for (ExcelTemplate excelTemplate : excelTemplates) {
//            Integer currentIndex = hasDealIndexMap.get(excelTemplate.sheetTitle()) == null ? excelTemplate.beginWriteRowIndex() : hasDealIndexMap.get(excelTemplate.sheetTitle());
//            sheet = workbook.getSheet(excelTemplate.sheetTitle());
//            row = sheet.createRow(currentIndex);
//           /* for (Map.Entry<Integer, Object> entry : excelTemplate.orderValueMap().entrySet()) {
//                row.createCell(entry.getKey()).setCellValue(entry.getValue() == null ? "" : entry.getValue().toString());
//            }*/
//            hasDealIndexMap.put(excelTemplate.sheetTitle(),++currentIndex);
//        }
//        return workbook;
//    }
//
//    public void outToFile() throws IOException {
//        Closer closer = Closer.create();
//        BufferedOutputStream out = closer.register(new BufferedOutputStream(new FileOutputStream(new File(outPath))));
//        this.putToWorkbook().write(out);
//        closer.close();
//    }
//
//    public static void main(String[] args) throws Exception {
//        new ExcelExportTest().outToFile();
//    }
//
//
//}
