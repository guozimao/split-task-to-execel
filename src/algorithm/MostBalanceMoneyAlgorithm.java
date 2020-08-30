package algorithm;

import beans.TaskExcel;
import utils.AlgorithmUtil;

import java.math.BigDecimal;
import java.util.*;

public class MostBalanceMoneyAlgorithm {

    public static void doProcess(List<TaskExcel> taskExelList,
                                  int baseNum,
                                  int counter,
                                  List<List<TaskExcel>> exportList,
                                  Boolean disableHistoryTakeNo,
                                  List<List<String>> taskNoHistoryList) {
        //存放小组的容器
        List<TaskExcel> subList = null;
        System.out.println("正在匹配并分组数据中...");
        //排序
        Collections.sort(taskExelList, new Comparator<TaskExcel>() {
            @Override
            public int compare(TaskExcel o1, TaskExcel o2) {
                BigDecimal left = o1.getPrice();
                BigDecimal right = o2.getPrice();
                return left.compareTo(right);
            }
        });
        //按分组算法逐一的匹配
        while(taskExelList.size() > 0){
            ListIterator<TaskExcel> listIterator = null;
            for (int i = 1; i <= baseNum; i++){
                //首个元素
                if(i == 1){
                    subList = new ArrayList<>();
                    listIterator = taskExelList.listIterator();
                    TaskExcel taskExcel = listIterator.next();
                    subList.add(taskExcel);
                    listIterator.remove();
                    counter ++;
                    //如果当前的数据就是最后一个
                    if(taskExelList.size() == 0){
                        exportList.add(subList);
                        counter = 1;
                        break;
                    }
                    //最后一个元素
                }else if(i == baseNum){
                    listIterator = taskExelList.listIterator();
                    //偶数逆向遍历
                    if(baseNum % 2 == 0){
                        while (listIterator.hasNext()){
                            listIterator.next();
                        }
                        while (listIterator.hasPrevious()){
                            TaskExcel taskExcel = listIterator.previous();
                            //小组中的任务编号不能重复
                            if(AlgorithmUtil.isSameTaskNo(taskExcel, subList)){
                                continue;
                            }else if(disableHistoryTakeNo && AlgorithmUtil.isExistHistoryTaskNo(taskExcel, subList,baseNum,taskNoHistoryList)){
                                continue;
                            }
                            subList.add(taskExcel);
                            listIterator.remove();
                            break;
                        }
                    }else{ //奇数正向遍历
                        while (listIterator.hasNext()){
                            TaskExcel taskExcel = listIterator.next();
                            //小组中的任务编号不能重复
                            if(AlgorithmUtil.isSameTaskNo(taskExcel, subList)) {
                                continue;
                            }else if(disableHistoryTakeNo && AlgorithmUtil.isExistHistoryTaskNo(taskExcel,subList,baseNum,taskNoHistoryList)){
                                continue;
                            }
                            subList.add(taskExcel);
                            listIterator.remove();
                            break;
                        }
                    }
                    exportList.add(subList);
                    AlgorithmUtil.setHistoryTaskNoList(subList,baseNum,taskNoHistoryList);
                    counter = 1;
                    // 偶数元素
                }else if(counter % 2 == 0){
                    listIterator = taskExelList.listIterator();
                    while (listIterator.hasNext()){
                        listIterator.next();
                    }
                    //逆向遍历
                    while (listIterator.hasPrevious()){
                        TaskExcel taskExcel = listIterator.previous();
                        //小组中的任务编号不能重复
                        if(AlgorithmUtil.isSameTaskNo(taskExcel, subList)){
                            continue;
                        }
                        subList.add(taskExcel);
                        listIterator.remove();
                        break;
                    }
                    counter ++;
                    //如果当前的数据就是最后一个
                    if(taskExelList.size() == 0){
                        exportList.add(subList);
                        counter = 1;
                        break;
                    }
                }else { // 奇数元素
                    listIterator = taskExelList.listIterator();
                    while (listIterator.hasNext()){
                        TaskExcel taskExcel = listIterator.next();
                        //小组中的任务编号不能重复
                        if(AlgorithmUtil.isSameTaskNo(taskExcel, subList)) {
                            continue;
                        }
                        subList.add(taskExcel);
                        listIterator.remove();
                        break;
                    }
                    counter ++;
                    //如果当前的数据就是最后一个
                    if(taskExelList.size() == 0){
                        exportList.add(subList);
                        counter = 1;
                        break;
                    }
                }
            }
        }
        System.out.println("匹配数据完毕...");
    }
}
