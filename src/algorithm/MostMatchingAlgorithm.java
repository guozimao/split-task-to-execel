package algorithm;

import beans.SortCount;
import beans.TaskExcel;
import org.apache.commons.collections4.Bag;
import org.apache.commons.collections4.MultiValuedMap;
import org.apache.commons.collections4.bag.HashBag;
import org.apache.commons.collections4.multimap.ArrayListValuedHashMap;
import utils.AlgorithmUtil;

import java.util.*;

public class MostMatchingAlgorithm {

    public static void doProcess(List<TaskExcel> taskExelList,
                                 int baseNum,
                                 List<List<TaskExcel>> exportList,
                                 boolean disableHistoryTakeNo,
                                 int counter,
                                 List<List<String>> taskNoHistoryList) {
        //存放小组的容器
        List<TaskExcel> subList = null;
        System.out.println("正在匹配并分组数据中...");
        sortTaskExelList4MostMatchingAlgorithm(taskExelList);
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
                    //如果当前的数据就是最后一个
                    if(taskExelList.size() == 0){
                        exportList.add(subList);
                        break;
                    }
                    //最后一个元素
                }else if(i == baseNum){
                    listIterator = taskExelList.listIterator();
                    while (listIterator.hasNext()){
                        TaskExcel taskExcel = listIterator.next();
                        //小组中的任务编号不能重复
                        if(AlgorithmUtil.isSameTaskNo(taskExcel, subList)) {
                            continue;
                        }else if(disableHistoryTakeNo && AlgorithmUtil.isExistHistoryTaskNo(taskExcel, subList,baseNum,taskNoHistoryList)){
                            continue;
                        }
                        subList.add(taskExcel);
                        listIterator.remove();
                        break;
                    }
                    exportList.add(subList);
                    AlgorithmUtil.setHistoryTaskNoList(subList,baseNum,taskNoHistoryList);
                }else {
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
                        break;
                    }
                }
            }
        }
        System.out.println("匹配数据完毕...");
    }

    private static void sortTaskExelList4MostMatchingAlgorithm(List<TaskExcel> taskExelList) {
        MultiValuedMap<String, TaskExcel> taskMap = new ArrayListValuedHashMap<>();
        Bag<String> taskBag = new HashBag<>();
        for(TaskExcel taskExcel : taskExelList){
            taskMap.put(taskExcel.getTaskNo(),taskExcel);
            taskBag.add(taskExcel.getTaskNo());
        }
        List<SortCount> sortingList = new ArrayList<>();
        for(String bag: taskBag.uniqueSet()){
            SortCount count = new SortCount();
            count.setCount(taskBag.getCount(bag));
            count.setTaskExcelList((List<TaskExcel>) taskMap.get(bag));
            sortingList.add(count);
        }
        //排序
        Collections.sort(sortingList, new Comparator<SortCount>() {
            @Override
            public int compare(SortCount o1, SortCount o2) {
                Integer count1 = o1.getCount();
                Integer count2 = o2.getCount();
                if(count1 == count2){
                    return 0;
                }else if(count1 > count2){
                    return -1;
                }else {
                    return 1;
                }
            }
        });
        List<TaskExcel> sortedList = new ArrayList<>();
        for(SortCount sortCount : sortingList){
            for(TaskExcel excel: sortCount.getTaskExcelList()){
                sortedList.add(excel);
            }
        }
        taskExelList = sortedList;
    }
}
