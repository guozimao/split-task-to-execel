import beans.MyPicture;
import beans.TaskExcel;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import utils.ImportExcel;

import java.io.*;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class Main {

    //任务列表
    private static List<TaskExcel> taskExelList = new LinkedList<>();
    //当前记录数
    private static int counter = 1;
    //分组的基数(默认基数为4)
    private static int baseNum = 4;
    //经算法处理后的列表
    private static List<List<TaskExcel>> exportList = new ArrayList<>();
    //未配对的其它列表
    private static List<TaskExcel> otherExportList = new ArrayList<>();
    //任务编号的历史列表
    private static List<List<String>> taskNoHistoryList = new ArrayList<>();

    public static void main(String[] args) {
        isOrNotVaild();
        setBaseNum(args);
        getExcelData();
        doProcessTask();
        exportIO();
        System.exit(0);
    }

    private static void isOrNotVaild() {
        LocalDate currentDate = LocalDate.now();
        if(currentDate.compareTo(LocalDate.of(2020, 11, 7 )) > 0){
            System.out.println("程序失效了");
            System.exit(0);
        }
    }

    /**
     * 批量导出excel表格
     *
     * */
    private static void exportIO() {
        //获取当前时间
        LocalDate currentDate = LocalDate.now();
        DateTimeFormatter dateTimeFormatter = DateTimeFormatter.ofPattern("yyyyMMdd");

        //创建文件目录
        String directory = "D:\\work-space\\" + currentDate.format(dateTimeFormatter);
        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }

        System.out.println("开始导出excel");
        for (int i = 0 ; i< exportList.size(); i++){
            if(exportList.get(i).size() != baseNum){
                exportList.get(i).stream().forEach( item -> otherExportList.add(item));
                continue;
            }
            exportMatchingExcel(i,currentDate,dateTimeFormatter);
        }
        exportNoMatchingExcel(currentDate,dateTimeFormatter);
        System.out.println("导出excel结束");

    }

    private static void exportNoMatchingExcel(LocalDate currentDate,DateTimeFormatter dateTimeFormatter) {

        if(otherExportList.size() == 0){
            return;
        }

        HSSFWorkbook workBook = new HSSFWorkbook();// 创建一个Excel工作薄
        HSSFSheet sheet = workBook.createSheet("sheet1");

        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();

        HSSFRow headerRow = sheet.createRow(0);// 创建首行，并赋值
        HSSFFont headFont = workBook.createFont();
        headFont.setFontName("仿宋_GB2312");
        headFont.setFontHeightInPoints((short) 14);
        headFont.setBold(true);
        HSSFCellStyle headerStyle = workBook.createCellStyle();
        headerStyle.setFont(headFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        for (int k = 0; k < ImportExcel.EXCEL_HEADER.length; k++) {//给首行赋值
            HSSFCell headerCell = headerRow.createCell(k);
            headerCell.setCellValue(ImportExcel.EXCEL_HEADER[k]);
            headerCell.setCellStyle(headerStyle);
            if(ImportExcel.EXCEL_HEADER[k].equals("主图") || ImportExcel.EXCEL_HEADER[k].equals("店铺名称（非掌柜名）")){
                sheet.setColumnWidth(k, 255 * 30);
            } else {
                sheet.setColumnWidth(k, 255 * 15);
            }
        }

        HSSFFont bodyFont = workBook.createFont();
        bodyFont.setFontName("仿宋_GB2312");
        bodyFont.setFontHeightInPoints((short) 12);
        bodyFont.setBold(true);

        //任务编号单元的样式
        HSSFCellStyle contentStyle4TaskNo =   workBook.createCellStyle();
        contentStyle4TaskNo.setAlignment(HorizontalAlignment.CENTER);
        contentStyle4TaskNo.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle4TaskNo.setFillForegroundColor(IndexedColors.RED.getIndex());
        contentStyle4TaskNo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        contentStyle4TaskNo.setBorderBottom(BorderStyle.DOUBLE); //底部边框
        contentStyle4TaskNo.setBottomBorderColor(IndexedColors.YELLOW.getIndex());//底部边框颜色
        contentStyle4TaskNo.setBorderLeft(BorderStyle.DOUBLE); //左边框
        contentStyle4TaskNo.setLeftBorderColor(IndexedColors.YELLOW.getIndex());//左边框颜色
        contentStyle4TaskNo.setBorderRight(BorderStyle.DOUBLE);
        contentStyle4TaskNo.setRightBorderColor(IndexedColors.YELLOW.getIndex());
        contentStyle4TaskNo.setBorderTop(BorderStyle.DOUBLE);
        contentStyle4TaskNo.setRightBorderColor(IndexedColors.YELLOW.getIndex());
        contentStyle4TaskNo.setFont(bodyFont);

        //备注样式
        HSSFCellStyle contentStyle4Note =   workBook.createCellStyle();
        contentStyle4Note.setAlignment(HorizontalAlignment.CENTER);
        contentStyle4Note.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle4Note.setWrapText(true);
        HSSFFont noteFont = workBook.createFont();
        noteFont.setBold(true);
        noteFont.setFontHeightInPoints((short) 12);
        noteFont.setColor(IndexedColors.RED.getIndex());
        contentStyle4Note.setFont(noteFont);

        //日期单元样式
        HSSFCellStyle dateStyle =   workBook.createCellStyle();
        HSSFCreationHelper createHelper = workBook.getCreationHelper();
        dateStyle.setWrapText(true);
        dateStyle.setAlignment(HorizontalAlignment.CENTER);
        dateStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/m/d"));

        //默认的输出单元样式
        HSSFCellStyle contentStyle =   workBook.createCellStyle();
        contentStyle.setWrapText(true);
        contentStyle.setAlignment(HorizontalAlignment.CENTER);
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle.setFont(bodyFont);

        for (int i = 0; i < otherExportList.size(); i++){
            HSSFRow row = sheet.createRow(i + 1);

            row.setHeight((short) 3219);
            HSSFCell hssfCell0 = row.createCell(0);
            hssfCell0.setCellValue(otherExportList.get(i).getTaskNo());
            hssfCell0.setCellStyle(contentStyle4TaskNo);

            HSSFCell hssfCell1 = row.createCell(1);
            hssfCell1.setCellValue(otherExportList.get(i).getDate());
            hssfCell1.setCellStyle(dateStyle);

            HSSFCell hssfCell3 = row.createCell(3);
            hssfCell3.setCellValue(otherExportList.get(i).getStoreName());
            hssfCell3.setCellStyle(contentStyle);

            HSSFCell hssfCell4 = row.createCell(4);
            hssfCell4.setCellValue(otherExportList.get(i).getPrice().doubleValue());
            hssfCell4.setCellStyle(contentStyle);

            HSSFCell hssfCell5 = row.createCell(5);
            hssfCell5.setCellValue(otherExportList.get(i).getNote());
            hssfCell5.setCellStyle(contentStyle4Note);

            HSSFCell hssfCell6 = row.createCell(6);
            hssfCell6.setCellValue(otherExportList.get(i).getSpecialNote());
            hssfCell6.setCellStyle(contentStyle4Note);

            HSSFCell hssfCell7 = row.createCell(7);
            hssfCell7.setCellValue(otherExportList.get(i).getKeyWord1());
            hssfCell7.setCellStyle(contentStyle);

            HSSFCell hssfCell8 = row.createCell(8);
            hssfCell8.setCellValue(otherExportList.get(i).getKeyWord2());
            hssfCell8.setCellStyle(contentStyle);

            if(otherExportList.get(i).getMyPicture() != null){
                //图片处理
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0,
                        (short) 2, i + 1, (short) 3, i + 2);
                MyPicture myPicture = otherExportList.get(i).getMyPicture();
                myPicture.setClientAnchor(anchor);

                // 插入图片
                patriarch.createPicture(anchor, workBook.addPicture(myPicture.getPictureData().getData(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            }

        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();// 将Excel文件存在输出流中
        try {
            workBook.write(os);// 将Excel写入输出流中
            byte[] fileContent = os.toByteArray();// 将输出流转换成字节数组
            os.close();
            OutputStream out = new FileOutputStream("D:\\work-space\\" + currentDate.format(dateTimeFormatter) + "\\" + "index.xls");
            out.write(fileContent);
            out.close();
            workBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private static void exportMatchingExcel(int i,LocalDate currentDate,DateTimeFormatter dateTimeFormatter) {
        String taskNoGroup = null;
        BigDecimal total = new BigDecimal(0);
        HSSFWorkbook workBook = new HSSFWorkbook();// 创建一个Excel工作薄
        HSSFSheet sheet = workBook.createSheet("sheet1");

        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();

        HSSFRow headerRow = sheet.createRow(0);// 创建首行，并赋值
        HSSFFont headFont = workBook.createFont();
        headFont.setFontName("仿宋_GB2312");
        headFont.setFontHeightInPoints((short) 14);
        headFont.setBold(true);
        HSSFCellStyle headerStyle = workBook.createCellStyle();
        headerStyle.setFont(headFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        for (int k = 0; k < ImportExcel.EXCEL_HEADER.length; k++) {//给首行赋值
            HSSFCell headerCell = headerRow.createCell(k);
            headerCell.setCellValue(ImportExcel.EXCEL_HEADER[k]);
            headerCell.setCellStyle(headerStyle);
            if(ImportExcel.EXCEL_HEADER[k].equals("主图") || ImportExcel.EXCEL_HEADER[k].equals("店铺名称（非掌柜名）")){
                sheet.setColumnWidth(k, 255 * 30);
            } else {
                sheet.setColumnWidth(k, 255 * 15);
            }
        }

        HSSFFont bodyFont = workBook.createFont();
        bodyFont.setFontName("仿宋_GB2312");
        bodyFont.setFontHeightInPoints((short) 12);
        bodyFont.setBold(true);

        //任务编号单元的样式
        HSSFCellStyle contentStyle4TaskNo =   workBook.createCellStyle();
        contentStyle4TaskNo.setAlignment(HorizontalAlignment.CENTER);
        contentStyle4TaskNo.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle4TaskNo.setFillForegroundColor(IndexedColors.RED.getIndex());
        contentStyle4TaskNo.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        contentStyle4TaskNo.setBorderBottom(BorderStyle.DOUBLE); //底部边框
        contentStyle4TaskNo.setBottomBorderColor(IndexedColors.YELLOW.getIndex());//底部边框颜色
        contentStyle4TaskNo.setBorderLeft(BorderStyle.DOUBLE); //左边框
        contentStyle4TaskNo.setLeftBorderColor(IndexedColors.YELLOW.getIndex());//左边框颜色
        contentStyle4TaskNo.setBorderRight(BorderStyle.DOUBLE);
        contentStyle4TaskNo.setRightBorderColor(IndexedColors.YELLOW.getIndex());
        contentStyle4TaskNo.setBorderTop(BorderStyle.DOUBLE);
        contentStyle4TaskNo.setRightBorderColor(IndexedColors.YELLOW.getIndex());
        contentStyle4TaskNo.setFont(bodyFont);

        //备注样式
        HSSFCellStyle contentStyle4Note =   workBook.createCellStyle();
        contentStyle4Note.setAlignment(HorizontalAlignment.CENTER);
        contentStyle4Note.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle4Note.setWrapText(true);
        HSSFFont noteFont = workBook.createFont();
        noteFont.setBold(true);
        noteFont.setFontHeightInPoints((short) 12);
        noteFont.setColor(IndexedColors.RED.getIndex());
        contentStyle4Note.setFont(noteFont);

        //日期单元样式
        HSSFCellStyle dateStyle =   workBook.createCellStyle();
        HSSFCreationHelper createHelper = workBook.getCreationHelper();
        dateStyle.setWrapText(true);
        dateStyle.setAlignment(HorizontalAlignment.CENTER);
        dateStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        dateStyle.setDataFormat(createHelper.createDataFormat().getFormat("yyyy/m/d"));

        //默认的输出单元样式
        HSSFCellStyle contentStyle =   workBook.createCellStyle();
        contentStyle.setWrapText(true);
        contentStyle.setAlignment(HorizontalAlignment.CENTER);
        contentStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        contentStyle.setFont(bodyFont);

        for (int j = 0; j < exportList.get(i).size(); j++){
            HSSFRow row = sheet.createRow(j + 1);

            row.setHeight((short) 3219);
            HSSFCell hssfCell0 = row.createCell(0);
            hssfCell0.setCellValue(exportList.get(i).get(j).getTaskNo());
            hssfCell0.setCellStyle(contentStyle4TaskNo);

            HSSFCell hssfCell1 = row.createCell(1);
            hssfCell1.setCellValue(exportList.get(i).get(j).getDate());
            hssfCell1.setCellStyle(dateStyle);

            HSSFCell hssfCell3 = row.createCell(3);
            hssfCell3.setCellValue(exportList.get(i).get(j).getStoreName());
            hssfCell3.setCellStyle(contentStyle);

            HSSFCell hssfCell4 = row.createCell(4);
            hssfCell4.setCellValue(exportList.get(i).get(j).getPrice().doubleValue());
            hssfCell4.setCellStyle(contentStyle);

            HSSFCell hssfCell5 = row.createCell(5);
            hssfCell5.setCellValue(exportList.get(i).get(j).getNote());
            hssfCell5.setCellStyle(contentStyle4Note);

            HSSFCell hssfCell6 = row.createCell(6);
            hssfCell6.setCellValue(exportList.get(i).get(j).getSpecialNote());
            hssfCell6.setCellStyle(contentStyle4Note);

            HSSFCell hssfCell7 = row.createCell(7);
            hssfCell7.setCellValue(exportList.get(i).get(j).getKeyWord1());
            hssfCell7.setCellStyle(contentStyle);

            HSSFCell hssfCell8 = row.createCell(8);
            hssfCell8.setCellValue(exportList.get(i).get(j).getKeyWord2());
            hssfCell8.setCellStyle(contentStyle);

            if(exportList.get(i).get(j).getMyPicture() != null){
                //图片处理
                HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 0, 0,
                        (short) 2, j + 1, (short) 3, j + 2);
                MyPicture myPicture = exportList.get(i).get(j).getMyPicture();
                myPicture.setClientAnchor(anchor);

                // 插入图片
                patriarch.createPicture(anchor, workBook.addPicture(myPicture.getPictureData().getData(), HSSFWorkbook.PICTURE_TYPE_JPEG));
            }

            if(j == 0){
                taskNoGroup = exportList.get(i).get(j).getTaskNo();
            }else {
                taskNoGroup = taskNoGroup + "-" + exportList.get(i).get(j).getTaskNo();
            }


            total = total.add(exportList.get(i).get(j).getPrice());

        }
        ByteArrayOutputStream os = new ByteArrayOutputStream();// 将Excel文件存在输出流中
        try {
            workBook.write(os);// 将Excel写入输出流中
            byte[] fileContent = os.toByteArray();// 将输出流转换成字节数组
            os.close();
            OutputStream out = new FileOutputStream("D:\\work-space\\" + currentDate.format(dateTimeFormatter) + "\\" + (i+1) + taskNoGroup + "-" + total.toPlainString() + "-" + currentDate.format(DateTimeFormatter.ofPattern("MMdd")) +".xls");
            out.write(fileContent);
            out.close();
            workBook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 分组算法（n个为一组,n为分组基数,并且TaskNo不能一样，而且每组之间的价格尽可能接近）
     *
     * **/
    private static void doProcessTask() {
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
                if(i == 1){ //首个元素
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
                }else if(i == baseNum){ //最后一个元素
                    listIterator = taskExelList.listIterator();
                    if(baseNum % 2 == 0){ //偶数逆向遍历
                        while (listIterator.hasNext()){
                            listIterator.next();
                        }
                        while (listIterator.hasPrevious()){
                            TaskExcel taskExcel = listIterator.previous();
                            //小组中的任务编号不能重复
                            if(isSameTaskNo(taskExcel, subList)){
                                continue;
                            }else if(isExistHistoryTaskNo(taskExcel, subList)){
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
                            if(isSameTaskNo(taskExcel, subList)) {
                                continue;
                            }else if(isExistHistoryTaskNo(taskExcel, subList)){
                                continue;
                            }
                            subList.add(taskExcel);
                            listIterator.remove();
                            break;
                        }
                    }
                    exportList.add(subList);
                    setHistoryTaskNoList(subList);
                    counter = 1;
                }else if(counter % 2 == 0){ // 偶数元素
                    listIterator = taskExelList.listIterator();
                    while (listIterator.hasNext()){
                        listIterator.next();
                    }
                    while (listIterator.hasPrevious()){ //逆向遍历
                        TaskExcel taskExcel = listIterator.previous();
                        //小组中的任务编号不能重复
                        if(isSameTaskNo(taskExcel, subList)){
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
                        if(isSameTaskNo(taskExcel, subList)) {
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

    /**
     * 判断历史存在的小组编号
     *
     * */
    private static boolean isExistHistoryTaskNo(TaskExcel taskExcel1, List<TaskExcel> subList) {
        //对于小于基数-1的小组不校验
        if(subList.size() < (baseNum - 1)) {
            return false;
        }
        //元素相同的个数
        int sameNum = 0;
        for(List<String> taskNOList : taskNoHistoryList){
            for(TaskExcel taskExcel : subList){ //已经添加的元素
                if(taskNOList.indexOf(taskExcel.getTaskNo()) > -1){
                    sameNum ++;
                }
            }
            if(taskNOList.indexOf(taskExcel1.getTaskNo()) > -1){ //预添加的元素
                sameNum ++;
            }
            if(sameNum == baseNum){
                return true;
            }
            sameNum = 0;
        }
        return false;
    }

    /**
     * 添加历史分组taskNO记录
     *
     * */
    private static void setHistoryTaskNoList(List<TaskExcel> subList) {
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
     * 判断是否有相同的taskNo
     *
     * **/
    private static boolean isSameTaskNo(TaskExcel taskExcel, List<TaskExcel> subList) {
        for(TaskExcel taskExcel1 : subList){
            if(taskExcel1.getTaskNo().equals(taskExcel.getTaskNo())){
                return true;
            }
        }
        return false;
    }

    /**
     * 将excel中的数据读到taskExcelList中
     *
     * **/
    private static void getExcelData() {
        Workbook wb =null;
        Sheet sheet = null;
        Row row = null;
        Object cellData = null;
        //创建文件目录
        String directory = "D:\\work-space\\";
        File file = new File(directory);
        if (!file.exists()) {
            file.mkdirs();
        }
        String columns[] = ImportExcel.EXCEL_HEADER;

        String filePath = "D:\\work-space\\index.xls";
        wb = ImportExcel.readExcel(filePath);

        System.out.println("正读取excel文件："+ filePath);

        if(wb == null){
            System.out.println("读取不到excel文件："+ filePath);
            filePath = "D:\\work-space\\index.xlsx";
            wb = ImportExcel.readExcel(filePath);
            System.out.println("正读取excel文件："+ filePath);
        }

        if(wb != null){
            //获取第一个sheet
            sheet = wb.getSheetAt(0);
            //获取最大行数
            int rownum = sheet.getPhysicalNumberOfRows();
            //获取第一行
            row = sheet.getRow(0);
            //获取最大列数
            int colnum = row.getPhysicalNumberOfCells();
            for (int i = 1; i<rownum; i++) {
                TaskExcel task = new TaskExcel();
                row = sheet.getRow(i);
                if(row != null && row.getPhysicalNumberOfCells() > 0){
                    Boolean isEmptyRow = false;
                    task.setRowNum(row.getRowNum());
                    for (int j=0;j<colnum;j++){
                        cellData = ImportExcel.getCellFormatValue(row.getCell(j));
                        if(columns[j].equals("任务代码")){
                            //过滤掉空行数据
                            if(cellData.equals("")){
                                isEmptyRow = true;
                                break;
                            }
                            task.setTaskNo((String)cellData);
                        }else if(columns[j].equals("日期")){
                            task.setDate((Date)cellData);
                        }else if(columns[j].equals("店铺名称（非掌柜名）")){
                            task.setStoreName((String)cellData);
                        }else if(columns[j].equals("单价/元")){
                            task.setPrice(new BigDecimal((String)cellData));
                        }else if(columns[j].equals("单价备注")){
                            task.setNote((String)cellData);
                        }else if(columns[j].equals("特殊备注")){
                            task.setSpecialNote((String)cellData);
                        }else if(columns[j].equals("关键词1")){
                            task.setKeyWord1((String)cellData);
                        }else if(columns[j].equals("关键词2")){
                            task.setKeyWord2((String)cellData);
                        }
                    }
                    if(!isEmptyRow){
                        taskExelList.add(task);
                    }
                }else{
                    break;
                }
            }
            try {
                if(ImportExcel.getIs03Or07()){
                    ImportExcel.getHSSFPictures((HSSFSheet) sheet, taskExelList);
                }else{
                    ImportExcel.getXSSFPictures((XSSFSheet) sheet, taskExelList);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            System.out.println("读取excel完毕...");
        }else {
            System.out.println("读取不到excel文件："+ filePath);
        }
    }
    /**
     * 设置基数
     *
     * **/
    private static void setBaseNum(String[] args){
        if(args != null && args.length > 0){
            Main.baseNum = Integer.valueOf(args[0]);
        }
    }
}
