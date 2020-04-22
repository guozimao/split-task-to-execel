package utils;

import beans.MyPicture;
import beans.TaskExcel;
import org.apache.poi.POIXMLDocumentPart;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTMarker;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;


public class ImportExcel {

    public static final String[] EXCEL_HEADER = new String[]{
            "任务代码",
            "日期",
            "主图",
            "店铺名称（非掌柜名）",
            "单价/元",
            "单价备注",
            "特殊备注",
            "关键词1",
            "关键词2"
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
            e.printStackTrace();
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
                for(TaskExcel task : taskList){
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
                    for(TaskExcel task : taskList){
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
}
