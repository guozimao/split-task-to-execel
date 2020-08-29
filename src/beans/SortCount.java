package beans;


import java.util.List;

public class SortCount {
    private Integer count;
    private List<TaskExcel> taskExcelList;

    public int getCount() {
        return count;
    }

    public void setCount(int count) {
        this.count = count;
    }

    public List<TaskExcel> getTaskExcelList() {
        return taskExcelList;
    }

    public void setTaskExcelList(List<TaskExcel> taskExcelList) {
        this.taskExcelList = taskExcelList;
    }
}
