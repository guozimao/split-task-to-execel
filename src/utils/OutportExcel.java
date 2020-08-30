package utils;


import beans.TaskExcel;

import scene.Exhibition2Salesman;

import java.io.*;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.List;

public class OutportExcel {

    /**
     * 批量导出excel表格
     *
     * */
    public static void exportIO4Salesman(Boolean disableCompositeFile4UnmatchedData,
                                 List<List<TaskExcel>> exportList,
                                 boolean enableNoMatchedMod,
                                 boolean disableHistoryTakeNo,
                                 int baseNum,
                                 List<TaskExcel> otherExportList) {
        //获取当前时间
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMdd");
        int order = 0;

        //创建文件目录
        String directory = null;
        if(disableCompositeFile4UnmatchedData){
            directory = "D:\\work-space\\" + currentDate.format(dateTimeFormatter);
        }else{
            directory = "D:\\work-space\\" + currentDate.format(dateTimeFormatter) + "\\data";
        }
        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }

        System.out.println("开始导出excel");
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
}
