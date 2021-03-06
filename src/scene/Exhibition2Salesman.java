package scene;

import beans.MyPicture;
import beans.TaskExcel;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import utils.OutportExcel;

import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class Exhibition2Salesman {

    public static void exportMatchingExcel(int i, int order, LocalDateTime currentDateTime, DateTimeFormatter dateTimeFormatter, List<List<TaskExcel>> exportList, boolean disableCompositeFile4UnmatchedData) {
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
        for (int k = 0; k < OutportExcel.SALESMAN_EXCEL_HEADER.length; k++) {
            HSSFCell headerCell = headerRow.createCell(k);
            headerCell.setCellValue(OutportExcel.SALESMAN_EXCEL_HEADER[k]);
            headerCell.setCellStyle(headerStyle);
            if(OutportExcel.SALESMAN_EXCEL_HEADER[k].equals("主图") || OutportExcel.SALESMAN_EXCEL_HEADER[k].equals("店铺名称（非掌柜名）")){
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

            HSSFCell hssfCell3 = row.createCell(3);
            hssfCell3.setCellValue(exportList.get(i).get(j).getStoreName());
            hssfCell3.setCellStyle(contentStyle);

            HSSFCell hssfCell4 = row.createCell(4);
            hssfCell4.setCellValue(exportList.get(i).get(j).getCallCenter());
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


            if(exportList.get(i).get(j).getMyPicture() != null){
                //图片处理
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0,
                        (short) 2, j + 1, (short) 3, j + 2);
                MyPicture myPicture = exportList.get(i).get(j).getMyPicture();
                myPicture.setClientAnchor(anchor);

                // 插入图片
                patriarch.createPicture(anchor, workBook.addPicture(myPicture.getPictureData().getData(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            }

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
            OutputStream out = null;
            if(disableCompositeFile4UnmatchedData){
                out = new FileOutputStream("D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\salesman\\" + (order+1) + taskNoGroup + "-" + total.toPlainString() + "-" + currentDateTime.format(DateTimeFormatter.ofPattern("MMdd")) +".xls");
            }else {
                out = new FileOutputStream("D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\salesman\\data\\" + (order+1) + taskNoGroup + "-" + total.toPlainString() + "-" + currentDateTime.format(DateTimeFormatter.ofPattern("MMdd")) +".xls");
            }
            out.write(fileContent);
            out.close();
            workBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void exportNoMatchingExcel(int i, int order,LocalDateTime currentDateTime,DateTimeFormatter dateTimeFormatter,List<TaskExcel> otherExportList) {

        String taskNoGroup = otherExportList.get(i).getTaskNo();
        BigDecimal total =  otherExportList.get(i).getPrice();
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
        for (int k = 0; k < OutportExcel.SALESMAN_EXCEL_HEADER.length; k++) {
            HSSFCell headerCell = headerRow.createCell(k);
            headerCell.setCellValue(OutportExcel.SALESMAN_EXCEL_HEADER[k]);
            headerCell.setCellStyle(headerStyle);
            if(OutportExcel.SALESMAN_EXCEL_HEADER[k].equals("主图") || OutportExcel.SALESMAN_EXCEL_HEADER[k].equals("店铺名称（非掌柜名）")){
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


        HSSFRow row = sheet.createRow(i + 1);

        row.setHeight((short) 3219);
        HSSFCell hssfCell0 = row.createCell(0);
        hssfCell0.setCellValue(otherExportList.get(i).getTaskNo());
        hssfCell0.setCellStyle(contentStyle4TaskNo);

        HSSFCell hssfCell1 = row.createCell(1);
        hssfCell1.setCellValue(otherExportList.get(i).getDate());
        hssfCell1.setCellStyle(dateStyle);

        HSSFCell hssfCell3 = row.createCell(3);
        hssfCell3.setCellValue(otherExportList.get(i).getStoreName());
        hssfCell3.setCellStyle(contentStyle);

        HSSFCell hssfCell4 = row.createCell(4);
        hssfCell4.setCellValue(otherExportList.get(i).getCallCenter());
        hssfCell4.setCellStyle(contentStyle);

        HSSFCell hssfCell5 = row.createCell(5);
        hssfCell5.setCellValue(otherExportList.get(i).getPrice().doubleValue());
        hssfCell5.setCellStyle(contentStyle);

        HSSFCell hssfCell6 = row.createCell(6);
        hssfCell6.setCellValue(otherExportList.get(i).getNote());
        hssfCell6.setCellStyle(contentStyle4Note);

        HSSFCell hssfCell7 = row.createCell(7);
        hssfCell7.setCellValue(otherExportList.get(i).getSpecialNote());
        hssfCell7.setCellStyle(contentStyle4Note);

        HSSFCell hssfCell8 = row.createCell(8);
        hssfCell8.setCellValue(otherExportList.get(i).getKeyWord());
        hssfCell8.setCellStyle(contentStyle);


        if(otherExportList.get(i).getMyPicture() != null){
            //图片处理
            HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0,
                    (short) 2,  1, (short) 3,  2);
            MyPicture myPicture = otherExportList.get(i).getMyPicture();
            myPicture.setClientAnchor(anchor);

            // 插入图片
            patriarch.createPicture(anchor, workBook.addPicture(myPicture.getPictureData().getData(), HSSFWorkbook.PICTURE_TYPE_JPEG));
        }
        // 将Excel文件存在输出流中
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try {
            // 将Excel写入输出流中
            workBook.write(os);
            // 将输出流转换成字节数组
            byte[] fileContent = os.toByteArray();
            os.close();
            OutputStream out = new FileOutputStream("D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\salesman\\" + (order+1) + taskNoGroup + "-" + total.toPlainString() + "-" + currentDateTime.format(DateTimeFormatter.ofPattern("MMdd")) +".xls");
            out.write(fileContent);
            out.close();
            workBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    public static void compositeFile4NoMatching2Excel(List<TaskExcel> otherExportList,LocalDateTime currentDateTime, DateTimeFormatter dateTimeFormatter) {
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
        for (int k = 0; k < OutportExcel.SALESMAN_EXCEL_HEADER.length; k++) {
            HSSFCell headerCell = headerRow.createCell(k);
            headerCell.setCellValue(OutportExcel.SALESMAN_EXCEL_HEADER[k]);
            headerCell.setCellStyle(headerStyle);
            if(OutportExcel.SALESMAN_EXCEL_HEADER[k].equals("主图") || OutportExcel.SALESMAN_EXCEL_HEADER[k].equals("店铺名称（非掌柜名）")){
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

        for(int i=0;i < otherExportList.size(); i++){
            HSSFRow row = sheet.createRow(i + 1);

            row.setHeight((short) 3219);
            HSSFCell hssfCell0 = row.createCell(0);
            hssfCell0.setCellValue(otherExportList.get(i).getTaskNo());
            hssfCell0.setCellStyle(contentStyle4TaskNo);

            HSSFCell hssfCell1 = row.createCell(1);
            hssfCell1.setCellValue(otherExportList.get(i).getDate());
            hssfCell1.setCellStyle(dateStyle);

            HSSFCell hssfCell3 = row.createCell(3);
            hssfCell3.setCellValue(otherExportList.get(i).getStoreName());
            hssfCell3.setCellStyle(contentStyle);

            HSSFCell hssfCell4 = row.createCell(4);
            hssfCell4.setCellValue(otherExportList.get(i).getCallCenter());
            hssfCell4.setCellStyle(contentStyle);

            HSSFCell hssfCell5 = row.createCell(5);
            hssfCell5.setCellValue(otherExportList.get(i).getPrice().doubleValue());
            hssfCell5.setCellStyle(contentStyle);

            HSSFCell hssfCell6 = row.createCell(6);
            hssfCell6.setCellValue(otherExportList.get(i).getNote());
            hssfCell6.setCellStyle(contentStyle4Note);

            HSSFCell hssfCell7 = row.createCell(7);
            hssfCell7.setCellValue(otherExportList.get(i).getSpecialNote());
            hssfCell7.setCellStyle(contentStyle4Note);

            HSSFCell hssfCell8 = row.createCell(8);
            hssfCell8.setCellValue(otherExportList.get(i).getKeyWord());
            hssfCell8.setCellStyle(contentStyle);


            if(otherExportList.get(i).getMyPicture() != null){
                //图片处理
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0,
                        (short) 2,  i+1, (short) 3,  i+2);
                MyPicture myPicture = otherExportList.get(i).getMyPicture();
                myPicture.setClientAnchor(anchor);

                // 插入图片
                patriarch.createPicture(anchor, workBook.addPicture(myPicture.getPictureData().getData(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            }
        }
        // 将Excel文件存在输出流中
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        try {
            // 将Excel写入输出流中
            workBook.write(os);
            // 将输出流转换成字节数组
            byte[] fileContent = os.toByteArray();
            os.close();
            OutputStream out = new FileOutputStream("D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\salesman\\index.xls");
            out.write(fileContent);
            out.close();
            workBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
