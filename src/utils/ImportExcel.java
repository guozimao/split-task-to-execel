package utils;

import beans.MyPicture;
import beans.TaskExcel;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.*;
import java.math.BigDecimal;
import java.util.*;


public class ImportExcel {

    public static final String[] EXCEL_HEADER = new String[]{
            "任务代码",
            "日期",
            "主图",
            "店铺名称（非掌柜名）",
            "链接",
            "单价/元",
            "单价备注",
            "特殊备注",
            "关键词1"
    };

    private static Boolean is03Or07 = true;

    //读取excel
    public static Workbook readExcel(String filePath){
        Workbook wb = null;
        if(filePath==null){
            return null;
        }
        String extString = filePath.substring(filePath.lastIndexOf("."));
        InputStream is = null;
        try {
            is = new FileInputStream(filePath);
            if(".xls".equals(extString)){
                ImportExcel.is03Or07 = true;
            }else if(".xlsx".equals(extString)){
                ImportExcel.is03Or07 = false;
            }else{
                return wb;
            }
            return ImportExcel.is03Or07 ? new HSSFWorkbook(is) : new XSSFWorkbook(is);

        } catch (FileNotFoundException e) {
            System.out.println("-------------");
        } catch (IOException e) {
            e.printStackTrace();
        }
        return wb;
    }
    public static Object getCellFormatValue(Cell cell){
        Object cellValue = null;
        if(cell!=null){
            //判断cell类型
            switch(cell.getCellType()){
                case Cell.CELL_TYPE_NUMERIC:{
                    if(DateUtil.isCellDateFormatted(cell)){
                        Date d = (Date) cell.getDateCellValue();
                        cellValue = d;
                    } else {
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    }
                    break;
                }
                case Cell.CELL_TYPE_FORMULA:{
                        cellValue = String.valueOf(cell.getNumericCellValue());
                    break;
                }
                case Cell.CELL_TYPE_STRING:{
                    cellValue = cell.getRichStringCellValue().getString();
                    break;
                }
                case Cell.CELL_TYPE_BLANK:{
                    cellValue = "";
                    break;
                }
                case Cell.CELL_TYPE_ERROR:{
                    cellValue = "非法字符";
                    break;
                }
                default:
                    cellValue = "未知类型";
            }
        }else{
            cellValue = "";
        }
        return cellValue;
    }

    /**
     * 03版图片处理
     * 说明Excel中的图片不在单元格内，而是悬浮在单元格之前，采用如下方式读取，但要求图片必须放在某个单元格之内也不能压住边框，否则获取的行数会有重复的。
     *
     * **/
    public static void getHSSFPictures (HSSFSheet sheet, List<TaskExcel> taskList) throws IOException {
        List<HSSFShape> list = sheet.getDrawingPatriarch().getChildren();
        for (HSSFShape shape : list) {
            if (shape instanceof HSSFPicture) {
                MyPicture myPicture = new MyPicture();
                HSSFPicture picture = (HSSFPicture) shape;
                HSSFClientAnchor cAnchor = picture.getClientAnchor();
                myPicture.setPictureData(picture.getPictureData());

                Iterator<TaskExcel> iterator = taskList.iterator();
                while (iterator.hasNext()){
                    TaskExcel task = iterator.next();
                    if(task.getRowNum() == cAnchor.getRow1()){
                        task.setMyPicture(myPicture);
                        break;
                    }
                }
            }
        }

    }

    /**
     * 07版图片处理
     * 说明Excel中的图片不在单元格内，而是悬浮在单元格之前，采用如下方式读取，但要求图片必须放在某个单元格之内也不能压住边框，否则获取的行数会有重复的。
     *
     * **/
    public static void getXSSFPictures (XSSFSheet sheet,List<TaskExcel> taskList) throws IOException {
    List<POIXMLDocumentPart> list = sheet.getRelations();
        for (POIXMLDocumentPart part : list) {
            if (part instanceof XSSFDrawing) {
                XSSFDrawing drawing = (XSSFDrawing) part;
                List<XSSFShape> shapes = drawing.getShapes();
                for (XSSFShape shape : shapes) {
                    XSSFPicture picture = (XSSFPicture) shape;
                    MyPicture myPicture = new MyPicture();
                    myPicture.setPictureData(picture.getPictureData());
                    XSSFClientAnchor anchor = picture.getPreferredSize();
                    CTMarker marker = anchor.getFrom();

                    Iterator<TaskExcel> iterator = taskList.iterator();
                    while (iterator.hasNext()){
                        TaskExcel task = iterator.next();
                        if(marker.getRow() == task.getRowNum()){
                            task.setMyPicture(myPicture);
                            break;
                        }
                    }
                }
            }
        }
    }

    public static Boolean getIs03Or07() {
        return is03Or07;
    }

    public static void setIs03Or07(Boolean is03Or07) {
        ImportExcel.is03Or07 = is03Or07;
    }

    /**
     * 将excel中的数据读到taskExcelList中
     *
     * **/
    public static void getExcelData(List<TaskExcel> taskExelList) {
        Workbook wb =null;
        Sheet sheet = null;
        Row row = null;
        Object cellData = null;
        //创建文件目录
        String directory = "D:\\work-space\\";
        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }
        String columns[] = ImportExcel.EXCEL_HEADER;

        String filePath = "D:\\work-space\\index.xls";
        wb = ImportExcel.readExcel(filePath);

        System.out.println("正读取excel文件："+ filePath);

        if(wb == null){
            System.out.println("读取不到excel文件："+ filePath);
            filePath = "D:\\work-space\\index.xlsx";
            wb = ImportExcel.readExcel(filePath);
            System.out.println("正读取excel文件："+ filePath);
        }

        if(wb != null){
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            //获取第一行
            row = sheet.getRow(0);
            //获取最大列数
            int colnum = row.getPhysicalNumberOfCells();
            for (int i = 1; i<rownum; i++) {
                TaskExcel task = new TaskExcel();
                row = sheet.getRow(i);
                if(row != null && row.getPhysicalNumberOfCells() > 0){
                    Boolean isEmptyRow = false;
                    task.setRowNum(row.getRowNum());
                    for (int j=0;j<colnum;j++){
                        cellData = ImportExcel.getCellFormatValue(row.getCell(j));
                        if(columns[j].equals("任务代码")){
                            //过滤掉空行数据
                            if(cellData.equals("")){
                                isEmptyRow = true;
                                break;
                            }
                            task.setTaskNo((String)cellData);
                        }else if(columns[j].equals("日期")){
                            task.setDate((Date)cellData);
                        }else if(columns[j].equals("店铺名称（非掌柜名）")){
                            task.setStoreName((String)cellData);
                        }else if(columns[j].equals("链接")){
                            task.setPlatformUrl((String)cellData);
                        } else if(columns[j].equals("单价/元")){
                            task.setPrice(new BigDecimal((String)cellData));
                        }else if(columns[j].equals("单价备注")){
                            task.setNote((String)cellData);
                        }else if(columns[j].equals("特殊备注")){
                            task.setSpecialNote((String)cellData);
                        }else if(columns[j].equals("关键词1")){
                            task.setKeyWord((String)cellData);
                        }
                    }
                    if(!isEmptyRow){
                        taskExelList.add(task);
                    }
                }else{
                    break;
                }
            }
            try {
                if(ImportExcel.getIs03Or07()){
                    ImportExcel.getHSSFPictures((HSSFSheet) sheet, taskExelList);
                }else{
                    ImportExcel.getXSSFPictures((XSSFSheet) sheet, taskExelList);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            System.out.println("读取excel完毕...");
        }else {
            System.out.println("读取不到excel文件："+ filePath);
        }
    }
}
