package utils;

import beans.TaskExcel;

import java.util.ArrayList;
import java.util.List;

public class AlgorithmUtil {
    /**
     * 判断是否有相同的taskNo
     *
     * **/
    public static boolean isSameTaskNo(TaskExcel taskExcel, List<TaskExcel> subList) {
        for(TaskExcel taskExcel1 : subList){
            if(taskExcel1.getTaskNo().equals(taskExcel.getTaskNo())){
                return true;
            }
        }
        return false;
    }

    /**
     * 添加历史分组taskNO记录
     *
     * */
    public static void setHistoryTaskNoList(List<TaskExcel> subList,int baseNum,List<List<String>> taskNoHistoryList) {
        //对于小于基数的小组不存历史
        if(subList.size() < baseNum) {
            return;
        }
        List<String> taskNoList = new ArrayList<>();
        for(TaskExcel taskExcel: subList){
            taskNoList.add(taskExcel.getTaskNo());
        }
        taskNoHistoryList.add(taskNoList);
    }

    /**
     * 判断历史存在的小组编号
     *
     * */
    public static boolean isExistHistoryTaskNo(TaskExcel taskExcel1, List<TaskExcel> subList,int baseNum,List<List<String>> taskNoHistoryList) {
        //对于小于基数-1的小组不校验
        if(subList.size() < (baseNum - 1)) {
            return false;
        }
        //元素相同的个数
        int sameNum = 0;
        for(List<String> taskNOList : taskNoHistoryList){
            //已经添加的元素
            for(TaskExcel taskExcel : subList){
                if(taskNOList.indexOf(taskExcel.getTaskNo()) > -1){
                    sameNum ++;
                }
            }
            //预添加的元素
            if(taskNOList.indexOf(taskExcel1.getTaskNo()) > -1){
                sameNum ++;
            }
            if(sameNum == baseNum){
                return true;
            }
            sameNum = 0;
        }
        return false;
    }
}
