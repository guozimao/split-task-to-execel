import algorithm.MostBalanceMoneyAlgorithm;
import algorithm.MostMatchingAlgorithm;
import beans.MyPicture;
import beans.TaskExcel;

import org.apache.commons.lang3.StringUtils;
import utils.ImportExcel;
import utils.OSSClientUtil;
import utils.OutportExcel;


import java.net.URL;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.*;

public class Main {

    //任务列表
    private static List<TaskExcel> taskExelList = new LinkedList<>();
    //当前记录数
    private static Integer counter = 1;
    //分组的基数(默认基数为4)
    private static Integer baseNum = 4;
    //允许不允许takeNo组的历史重复
    private static Boolean disableHistoryTakeNo = true;
    //是否禁用对于未匹配进行合成文件
    private static Boolean disableCompositeFile4UnmatchedData = true;
    //是否禁用未匹配
    private static Boolean enableNoMatchedMod = true;
    //选择算法
    private static Integer algorithmIndex = 0;
    //补单
    private static Boolean disableSupplementOrderMode = true;
    //算法
    private static List<String> algorithm = Arrays.asList(new String[]{"mostBalanceMoney","mostMatching"});
    //经算法处理后的列表
    private static List<List<TaskExcel>> exportList = new ArrayList<>();
    //未配对的其它列表
    private static List<TaskExcel> otherExportList = new ArrayList<>();
    //任务编号的历史列表
    private static List<List<String>> taskNoHistoryList = new ArrayList<>();

    public static void main(String[] args) {
        isOrNotVaild();
        setProgameParam(args);
        ImportExcel.getExcelData(taskExelList);
        uploadPicture2Oss();
        doProcessTask();
        outportExcel();
        System.exit(0);
    }

    private static void outportExcel() {
        LocalDateTime currentDateTime = LocalDateTime.now();
        if(disableSupplementOrderMode){
            OutportExcel.exportIO4Salesman(disableCompositeFile4UnmatchedData,
                    exportList,enableNoMatchedMod,disableHistoryTakeNo,baseNum,otherExportList,currentDateTime);
            OutportExcel.exportIO4BackgroundStorage(exportList,enableNoMatchedMod,currentDateTime);
            OutportExcel.exportIO4PictureAndOssParam(exportList,currentDateTime);
        }else {
            OutportExcel.exportIO4PictureAndOssParam(exportList,currentDateTime);
            OutportExcel.exportIO4SupplementOrder(exportList,currentDateTime);
        }
    }

    private static void uploadPicture2Oss() {
        Map<String,MyPicture> pUrlMyPictureMap = new HashMap<>();
        for(TaskExcel taskExcel:taskExelList){
            pUrlMyPictureMap.put(taskExcel.getPlatformUrl(),taskExcel.getMyPicture());
        }
        pUrlMyPictureMap.remove("");
        pUrlMyPictureMap.remove(null);
        Map<String,String> pUrlOssPictureParam = new HashMap<>();
        System.out.println("开始上传图片到阿里Oss服务器");
        for(Map.Entry<String,MyPicture> excelEntry:pUrlMyPictureMap.entrySet()){
            String pictureName = UUID.randomUUID().toString() + ".png";
            URL pictureUrl = OSSClientUtil.picOSS(excelEntry.getValue().getPictureData().getData(),pictureName);
            String path = pictureUrl.getPath();
            String queryParam = pictureUrl.getQuery();
            String expires = StringUtils.substringBetween(queryParam + "&","Expires=","&");
            String signature = StringUtils.substringBetween(queryParam + "&","Signature=","&");
            String ossPictureParam = StringUtils.joinWith(",",path,expires,signature);
            pUrlOssPictureParam.put(excelEntry.getKey(),ossPictureParam);
            System.out.println("图片链接生成：" + pictureUrl.toString());
        }
        System.out.println("上传图片到阿里Oss服务器，已完成");
        for(TaskExcel taskExcel:taskExelList){
            taskExcel.setOssPictureParam(pUrlOssPictureParam.get(taskExcel.getPlatformUrl()));
        }
    }

    private static void isOrNotVaild() {
        LocalDate currentDate = LocalDate.now();
        if(currentDate.compareTo(LocalDate.of(2020, 11, 7 )) > 0){
            System.out.println("程序失效了");
            System.exit(0);
        }
    }

    /**
     * 分组算法（n个为一组,n为分组基数,并且TaskNo不能一样，而且每组之间的价格尽可能接近）
     *
     * **/
    private static void doProcessTask() {
        if(algorithm.indexOf("mostBalanceMoney") == Main.algorithmIndex){
            MostBalanceMoneyAlgorithm.doProcess(taskExelList,baseNum,counter,exportList,disableHistoryTakeNo,taskNoHistoryList);
        }else if(algorithm.indexOf("mostMatching") == Main.algorithmIndex){
            MostMatchingAlgorithm.doProcess(taskExelList,baseNum,exportList,disableHistoryTakeNo,counter,taskNoHistoryList);
        }else {
            System.out.println("没选择算法");
            System.exit(0);
        }
    }

    /**
     * 设置程序参数
     *
     * **/
    private static void setProgameParam(String[] args){
        //设置分组基数
        if(args != null && args.length > 0){
            Main.baseNum = Integer.valueOf(args[0]);
        }
        //设置分组成员允不允许重复
        if(args != null && args.length > 1){
            Main.disableHistoryTakeNo = Boolean.valueOf(args[1]);
        }
        //设置对于未匹配进行合成文件
        if(args != null && args.length > 2){
            Main.disableCompositeFile4UnmatchedData = Boolean.valueOf(args[2]);
        }
        //设置要不要未匹配
        if(args != null && args.length > 3){
            Main.enableNoMatchedMod = Boolean.valueOf(args[3]);
        }
        //设置算法
        if(args != null && args.length > 4){
            Main.algorithmIndex = Integer.valueOf(args[4]);
        }
        //设置是否处于补单模式
        if(args != null && args.length > 5){
            Main.disableSupplementOrderMode = Boolean.valueOf(args[5]);
        }
    }
}
