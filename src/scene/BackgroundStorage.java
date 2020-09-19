package scene;

import beans.TaskExcel;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import utils.OutportExcel;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class BackgroundStorage {
    public static void exportMatchingExcel(int i, int order, LocalDate currentDate, DateTimeFormatter dateTimeFormatter, List<List<TaskExcel>> exportList) {
        String taskNoGroup = null;
        BigDecimal total = new BigDecimal(0);
        // 创建一个Excel工作薄
        HSSFWorkbook workBook = new HSSFWorkbook();
        HSSFSheet sheet = workBook.createSheet("sheet1");

        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();

        // 创建首行，并赋值
        HSSFRow headerRow = sheet.createRow(0);
        HSSFFont headFont = workBook.createFont();
        headFont.setFontName("仿宋_GB2312");
        headFont.setFontHeightInPoints((short) 14);
        headFont.setBold(true);
        HSSFCellStyle headerStyle = workBook.createCellStyle();
        headerStyle.setFont(headFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        //给首行赋值
        for (int k = 0; k < OutportExcel.BACKGROUND_EXCEL_HEADER.length; k++) {
            HSSFCell headerCell = headerRow.createCell(k);
            headerCell.setCellValue(OutportExcel.BACKGROUND_EXCEL_HEADER[k]);
            headerCell.setCellStyle(headerStyle);
            if(OutportExcel.BACKGROUND_EXCEL_HEADER[k].equals("店铺名称（非掌柜名）")){
                sheet.setColumnWidth(k, 255 * 30);
            } else {
                sheet.setColumnWidth(k, 255 * 15);
            }
        }

        HSSFFont bodyFont = workBook.createFont();
        bodyFont.setFontName("仿宋_GB2312");
        bodyFont.setFontHeightInPoints((short) 12);
        bodyFont.setBold(true);

        //任务编号单元的样式
        HSSFCellStyle contentStyle4TaskNo =   workBook.createCellStyle();
        contentStyle4TaskNo.setAlignment(HorizontalAlignment.CENTER);
        contentStyle4TaskNo.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle4TaskNo.setFillForegroundColor(IndexedColors.RED.getIndex());
        contentStyle4TaskNo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        //底部边框
        contentStyle4TaskNo.setBorderBottom(BorderStyle.DOUBLE);
        //底部边框颜色
        contentStyle4TaskNo.setBottomBorderColor(IndexedColors.YELLOW.getIndex());
        //左边框
        contentStyle4TaskNo.setBorderLeft(BorderStyle.DOUBLE);
        //左边框颜色
        contentStyle4TaskNo.setLeftBorderColor(IndexedColors.YELLOW.getIndex());
        contentStyle4TaskNo.setBorderRight(BorderStyle.DOUBLE);
        contentStyle4TaskNo.setRightBorderColor(IndexedColors.YELLOW.getIndex());
        contentStyle4TaskNo.setBorderTop(BorderStyle.DOUBLE);
        contentStyle4TaskNo.setRightBorderColor(IndexedColors.YELLOW.getIndex());
        contentStyle4TaskNo.setFont(bodyFont);

        //备注样式
        HSSFCellStyle contentStyle4Note =   workBook.createCellStyle();
        contentStyle4Note.setAlignment(HorizontalAlignment.CENTER);
        contentStyle4Note.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle4Note.setWrapText(true);
        HSSFFont noteFont = workBook.createFont();
        noteFont.setBold(true);
        noteFont.setFontHeightInPoints((short) 12);
        noteFont.setColor(IndexedColors.RED.getIndex());
        contentStyle4Note.setFont(noteFont);

        //日期单元样式
        HSSFCellStyle dateStyle =   workBook.createCellStyle();
        HSSFCreationHelper createHelper = workBook.getCreationHelper();
        dateStyle.setWrapText(true);
        dateStyle.setAlignment(HorizontalAlignment.CENTER);
        dateStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/m/d"));

        //默认的输出单元样式
        HSSFCellStyle contentStyle =   workBook.createCellStyle();
        contentStyle.setWrapText(true);
        contentStyle.setAlignment(HorizontalAlignment.CENTER);
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle.setFont(bodyFont);

        for (int j = 0; j < exportList.get(i).size(); j++){
            HSSFRow row = sheet.createRow(j + 1);

            row.setHeight((short) 3219);
            HSSFCell hssfCell0 = row.createCell(0);
            hssfCell0.setCellValue(exportList.get(i).get(j).getTaskNo());
            hssfCell0.setCellStyle(contentStyle4TaskNo);

            HSSFCell hssfCell1 = row.createCell(1);
            hssfCell1.setCellValue(exportList.get(i).get(j).getDate());
            hssfCell1.setCellStyle(dateStyle);

            HSSFCell hssfCell2 = row.createCell(2);
            hssfCell2.setCellValue(exportList.get(i).get(j).getOssPictureParam());
            hssfCell2.setCellStyle(contentStyle);

            HSSFCell hssfCell3 = row.createCell(3);
            hssfCell3.setCellValue(exportList.get(i).get(j).getStoreName());
            hssfCell3.setCellStyle(contentStyle);

            HSSFCell hssfCell4 = row.createCell(4);
            hssfCell4.setCellValue(exportList.get(i).get(j).getPlatformUrl());
            hssfCell4.setCellStyle(contentStyle);

            HSSFCell hssfCell5 = row.createCell(5);
            hssfCell5.setCellValue(exportList.get(i).get(j).getPrice().doubleValue());
            hssfCell5.setCellStyle(contentStyle);

            HSSFCell hssfCell6 = row.createCell(6);
            hssfCell6.setCellValue(exportList.get(i).get(j).getNote());
            hssfCell6.setCellStyle(contentStyle4Note);

            HSSFCell hssfCell7 = row.createCell(7);
            hssfCell7.setCellValue(exportList.get(i).get(j).getSpecialNote());
            hssfCell7.setCellStyle(contentStyle4Note);

            HSSFCell hssfCell8 = row.createCell(8);
            hssfCell8.setCellValue(exportList.get(i).get(j).getKeyWord());
            hssfCell8.setCellStyle(contentStyle);


            if(j == 0){
                taskNoGroup = exportList.get(i).get(j).getTaskNo();
            }else {
                taskNoGroup = taskNoGroup + "-" + exportList.get(i).get(j).getTaskNo();
            }


            total = total.add(exportList.get(i).get(j).getPrice());

        }
        // 将Excel文件存在输出流中
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try {
            // 将Excel写入输出流中
            workBook.write(os);
            // 将输出流转换成字节数组
            byte[] fileContent = os.toByteArray();
            os.close();
            OutputStream out = new FileOutputStream("D:\\work-space\\" + currentDate.format(dateTimeFormatter) + "\\BackgroundStorage\\" + (order+1) + taskNoGroup + "-" + total.toPlainString() + "-" + currentDate.format(DateTimeFormatter.ofPattern("MMdd")) +".xls");
            out.write(fileContent);
            out.close();
            workBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
