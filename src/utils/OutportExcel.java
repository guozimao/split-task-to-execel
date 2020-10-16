package utils;


import beans.TaskExcel;

import org.apache.commons.collections4.MultiValuedMap;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import scene.BackgroundStorage;
import scene.BackgroundSupplementOrder;
import scene.ExHibition2Admin;
import scene.Exhibition2Salesman;

import java.io.*;

import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.ArrayList;
import java.util.List;

public class OutportExcel {

    public static final String[] BACKGROUND_EXCEL_HEADER = new String[]{
            "任务代码",
            "日期",
            "主图oss参数",
            "店铺名称（非掌柜名）",
            "客服",
            "链接",
            "单价/元",
            "单价备注",
            "特殊备注",
            "关键词"
    };

    public static final String[] SALESMAN_EXCEL_HEADER = new String[]{
            "任务代码",
            "日期",
            "主图",
            "店铺名称（非掌柜名）",
            "客服",
            "单价/元",
            "单价备注",
            "特殊备注",
            "关键词"
    };

    public static final String[] ADMIN_EXCEL_HEADER = new String[]{
            "任务代码",
            "日期",
            "主图",
            "主图oss参数",
            "店铺名称（非掌柜名）",
            "客服",
            "链接",
            "单价/元",
            "单价备注",
            "特殊备注",
            "关键词"
    };

    public static final String[] SUPPLEMENT_ORDER_EXCEL_HEADER = new String[]{
            "任务代码",
            "日期",
            "主图oss参数",
            "店铺名称（非掌柜名）",
            "客服",
            "链接",
            "单价/元",
            "单价备注",
            "特殊备注",
            "关键词"
    };

    /**
     * 批量导出excel表格
     *
     * */
    public static void exportIO4Salesman(Boolean disableCompositeFile4UnmatchedData,
                                 List<List<TaskExcel>> exportList,
                                 boolean enableNoMatchedMod,
                                 boolean disableHistoryTakeNo,
                                 Integer baseNum,
                                 List<TaskExcel> otherExportList,
                                         LocalDateTime currentDateTime) {
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        int order = 0;

        //创建文件目录
        String directory = null;
        if(disableCompositeFile4UnmatchedData){
            directory = "D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\salesman";
        }else{
            directory = "D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\salesman\\data";
        }
        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }

        System.out.println("开始导出Salesman的excel");
        // 匹配的组导出excel
        for (int i = 0; i< exportList.size(); i++){
            if(enableNoMatchedMod){
                if(disableHistoryTakeNo){
                    //在不重复算法的情况下，除了匹配数为baseNum，其它都是未匹配的组
                    if(exportList.get(i).size() != baseNum){
                        exportList.get(i).stream().forEach( item -> otherExportList.add(item));
                        continue;
                    }
                }else{
                    //在重复算法的情况下，只有匹配数为1的，算是未匹配的组
                    if(exportList.get(i).size() == 1){
                        exportList.get(i).stream().forEach( item -> otherExportList.add(item));
                        continue;
                    }
                }
            }
            Exhibition2Salesman.exportMatchingExcel(i,order,currentDateTime,dateTimeFormatter,exportList,disableCompositeFile4UnmatchedData);
            order ++;
        }
        // 未匹配的组导出excel
        if(enableNoMatchedMod){
            if(disableCompositeFile4UnmatchedData){
                for (int i = 0; i< otherExportList.size(); i++){
                    Exhibition2Salesman.exportNoMatchingExcel(i,order,currentDateTime,dateTimeFormatter,otherExportList);
                    order ++;
                }
            }else {
                Exhibition2Salesman.compositeFile4NoMatching2Excel(otherExportList,currentDateTime,dateTimeFormatter);
            }
        }

        System.out.println("导出excel结束");

    }

    /**
     * 批量导出excel表格，没考虑未匹配模式
     *
     * */
    public static void exportIO4BackgroundStorage(
                                         List<List<TaskExcel>> exportList,
                                         boolean enableNoMatchedMod,LocalDateTime currentDateTime) {
        if(enableNoMatchedMod){
            return;
        }

        //获取当前时间
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");
        int order = 0;

        //创建文件目录
        String directory = null;
        directory = "D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\BackgroundStorage";
        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }

        System.out.println("开始导出BackgroundStorage的excel");
        // 匹配的组导出excel
        for (int i = 0; i< exportList.size(); i++){
            BackgroundStorage.exportMatchingExcel(i,order,currentDateTime,dateTimeFormatter,exportList);
            order ++;
        }

        System.out.println("导出excel结束");

    }

    public static void exportIO4PictureAndOssParam(List<List<TaskExcel>> exportList,LocalDateTime currentDateTime) {
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

        //创建文件目录
        String directory = null;

        directory = "D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\admin";

        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }

        System.out.println("开始导出admin的excel");

        MultiValuedMap<String,TaskExcel> taskExcelsMap = new ArrayListValuedHashMap<>();
        for(List<TaskExcel> excelList : exportList){
            for(TaskExcel task : excelList){
                taskExcelsMap.put(task.getPlatformUrl(),task);
            }
        }

        List<TaskExcel> taskExcels = new ArrayList<>();
        for(String url : taskExcelsMap.keySet()){
            for(TaskExcel taskExcel : taskExcelsMap.get(url)){
                taskExcels.add(taskExcel);
            }
        }

        ExHibition2Admin.compositeFile4OssParam2Excel(taskExcels,currentDateTime,dateTimeFormatter);

        System.out.println("导出admin的excel结束");
    }

    public static void exportIO4SupplementOrder(List<List<TaskExcel>> exportList,LocalDateTime currentDateTime) {
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMddHHmmss");

        //创建文件目录
        String directory = null;

        directory = "D:\\work-space\\" + currentDateTime.format(dateTimeFormatter) + "\\supplementOrder";

        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }

        System.out.println("开始导出补单的excel");

        MultiValuedMap<String,TaskExcel> taskExcelsMap = new ArrayListValuedHashMap<>();
        for(List<TaskExcel> excelList : exportList){
            for(TaskExcel task : excelList){
                taskExcelsMap.put(task.getPlatformUrl(),task);
            }
        }

        List<TaskExcel> taskExcels = new ArrayList<>();
        for(String url : taskExcelsMap.keySet()){
            for(TaskExcel taskExcel : taskExcelsMap.get(url)){
                taskExcels.add(taskExcel);
            }
        }

        BackgroundSupplementOrder.compositeFile4SupplementOrder2Excel(taskExcels,currentDateTime,dateTimeFormatter);

        System.out.println("导出补单的excel结束");
    }

}
