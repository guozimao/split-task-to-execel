package utils;


import beans.TaskExcel;

import scene.BackgroundStorage;
import scene.ExHibition2Admin;
import scene.Exhibition2Salesman;

import java.io.*;

import java.time.LocalDate;
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

    /**
     * 批量导出excel表格
     *
     * */
    public static void exportIO4Salesman(Boolean disableCompositeFile4UnmatchedData,
                                 List<List<TaskExcel>> exportList,
                                 boolean enableNoMatchedMod,
                                 boolean disableHistoryTakeNo,
                                 Integer baseNum,
                                 List<TaskExcel> otherExportList) {
        //获取当前时间
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMdd");
        int order = 0;

        //创建文件目录
        String directory = null;
        if(disableCompositeFile4UnmatchedData){
            directory = "D:\\work-space\\" + currentDate.format(dateTimeFormatter) + "\\salesman";
        }else{
            directory = "D:\\work-space\\" + currentDate.format(dateTimeFormatter) + "\\salesman\\data";
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
            Exhibition2Salesman.exportMatchingExcel(i,order,currentDate,dateTimeFormatter,exportList,disableCompositeFile4UnmatchedData);
            order ++;
        }
        // 未匹配的组导出excel
        if(enableNoMatchedMod){
            if(disableCompositeFile4UnmatchedData){
                for (int i = 0; i< otherExportList.size(); i++){
                    Exhibition2Salesman.exportNoMatchingExcel(i,order,currentDate,dateTimeFormatter,otherExportList);
                    order ++;
                }
            }else {
                Exhibition2Salesman.compositeFile4NoMatching2Excel(otherExportList,currentDate,dateTimeFormatter);
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
                                         boolean enableNoMatchedMod) {
        if(enableNoMatchedMod){
            return;
        }

        //获取当前时间
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMdd");
        int order = 0;

        //创建文件目录
        String directory = null;
        directory = "D:\\work-space\\" + currentDate.format(dateTimeFormatter) + "\\BackgroundStorage";
        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }

        System.out.println("开始导出BackgroundStorage的excel");
        // 匹配的组导出excel
        for (int i = 0; i< exportList.size(); i++){
            BackgroundStorage.exportMatchingExcel(i,order,currentDate,dateTimeFormatter,exportList);
            order ++;
        }

        System.out.println("导出excel结束");

    }

    public static void exportIO4PictureAndOssParam(List<List<TaskExcel>> exportList) {
        //获取当前时间
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMdd");

        //创建文件目录
        String directory = null;

        directory = "D:\\work-space\\" + currentDate.format(dateTimeFormatter) + "\\admin";

        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }

        System.out.println("开始导出admin的excel");

        List<TaskExcel> taskExcels = new ArrayList<>();
        for(List<TaskExcel> excelList : exportList){
            for(TaskExcel task : excelList){
                taskExcels.add(task);
            }
        }

        ExHibition2Admin.compositeFile4OssParam2Excel(taskExcels,currentDate,dateTimeFormatter);

        System.out.println("导出admin的excel结束");
    }
}
