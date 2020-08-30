package beans;

import java.math.BigDecimal;
import java.util.Date;

public class TaskExcel {
    private int rowNum;
    private String taskNo;
    private Date date;
    private String storeName;
    private BigDecimal price;
    private String note;
    private String specialNote;
    private MyPicture myPicture;
    private String platformUrl;
    private String ossPictureParam;
    private String keyWord1;
    private String keyWord2;

    public String getSpecialNote() {
        return specialNote;
    }

    public void setSpecialNote(String specialNote) {
        this.specialNote = specialNote;
    }

    public int getRowNum() {
        return rowNum;
    }

    public void setRowNum(int rowNum) {
        this.rowNum = rowNum;
    }

    public String getTaskNo() {
        return taskNo;
    }

    public void setTaskNo(String taskNo) {
        this.taskNo = taskNo;
    }

    public Date getDate() {
        return date;
    }

    public void setDate(Date date) {
        this.date = date;
    }

    public String getStoreName() {
        return storeName;
    }

    public void setStoreName(String storeName) {
        this.storeName = storeName;
    }

    public BigDecimal getPrice() {
        return price;
    }

    public void setPrice(BigDecimal price) {
        this.price = price;
    }

    public String getNote() {
        return note;
    }

    public void setNote(String note) {
        this.note = note;
    }

    public MyPicture getMyPicture() {
        return myPicture;
    }

    public void setMyPicture(MyPicture myPicture) {
        this.myPicture = myPicture;
    }

    public String getKeyWord1() {
        return keyWord1;
    }

    public void setKeyWord1(String keyWord1) {
        this.keyWord1 = keyWord1;
    }

    public String getKeyWord2() {
        return keyWord2;
    }

    public void setKeyWord2(String keyWord2) {
        this.keyWord2 = keyWord2;
    }

    public String getPlatformUrl() {
        return platformUrl;
    }

    public void setPlatformUrl(String platformUrl) {
        this.platformUrl = platformUrl;
    }

    public String getOssPictureParam() {
        return ossPictureParam;
    }

    public void setOssPictureParam(String ossPictureParam) {
        this.ossPictureParam = ossPictureParam;
    }
}
