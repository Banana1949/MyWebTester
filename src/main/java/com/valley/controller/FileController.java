package com.valley.controller;


import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * @Description 此类示例基本的excel的导入和导出功能。依赖包：poi-3.17.jar、poi-ooxml-3.17.jar
 * @Author MuXin
 * @Date 2018/8/29 13:20
 * @Version 1.0
 **/
@Controller
@RequestMapping("file")
public class FileController{

    @RequestMapping("uploadExcel")
    @ResponseBody
    public String uploadExcel(@RequestParam("file") MultipartFile file,HttpServletResponse response){
        response.setCharacterEncoding("UTF-8");
        int checkFile = checkFile(file);
        if(checkFile == -1){
            return "file is empty!";
        }else if (checkFile == -2){
            return "not valid excel file!";
        }
        try {
            StringBuilder stringBuilder = new StringBuilder();
            //获取文件名
            String fileName = file.getOriginalFilename();
            stringBuilder.append("文件名："+fileName);
            Workbook workbook = getWorkBook(file);
            Sheet sheet=null;
            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {//循环每个sheet
                sheet=workbook.getSheetAt(i);//依次获得每个sheet
                String sheetName = sheet.getSheetName();//获得当前sheet名称
                stringBuilder.append("\n\n"+"第"+(i+1)+"个sheet名："+sheetName+"\n");
                for (int j = 0; j <= sheet.getLastRowNum(); j++) {//循环每一行
                    stringBuilder.append("第"+(j+1)+"行：");
                    Row row = sheet.getRow(j);
                    //如果在excel新增一个sheet却没有增加任何行时导致在新sheet中获取的row为null
                    if(row == null) {
                        continue;
                    }
                    for (int k = 0; k <= row.getLastCellNum(); k++){//循环每一个单元格
                        Cell cell = row.getCell(k);
                        String cellValue = getCellValueToString(cell);
                        stringBuilder.append("\t"+cellValue);
                    }
                    stringBuilder.append("\n");
                }
            }
            System.out.println(stringBuilder.toString());
            return "success!";
        }catch (Exception e){
            e.printStackTrace();
            return "unknown error!";
        }
    }


    @RequestMapping("downloadExcel")
    @ResponseBody
    public void downloadExcel(HttpServletRequest request, HttpServletResponse response) throws IOException {
        //模拟从持久层或其他数据源获取7w条数据
        List<User> users = new ArrayList<>();
        User user = null;
        for (int i=0;i<70000;i++){
            user = new User();
            user.setUserId(i);
            user.setUserName("Mr."+i);
            user.setBirthDate(new Date());
            users.add(user);
        }

        //创建excel
        String fileName = "用户数据"+new SimpleDateFormat("yyyyMMddHHmmss").format(new Date());
        //创建workbook
        Workbook work = null;
        if(users.size()<=65535){//xls格式的excel文件每个sheet支持的最大行数为65536行，此处多留出一行存放表头
            work = new HSSFWorkbook();
            fileName += ".xls";
        }else if (users.size()<=1048576) {//xlsx格式的excel文件每个sheet支持的最大行数为1048576行，此处多留出一行存放表头
            work = new XSSFWorkbook();
            fileName += ".xlsx";
        }else {//超出目前单个sheet所能容纳行数上限。此处仅仅返回空excel，也可以根据情况分sheet导出。
            work = new HSSFWorkbook();
            fileName = "数据太多啦.xls";
            users.clear();
        }
        //创建sheet
        Sheet sheet = work.createSheet("工作表1");
        //添加表头行,并增加表头行单元格内容
        Row titleRow = sheet.createRow(0);
        Cell titleCell0 = titleRow.createCell(0);
        titleCell0.setCellValue("ID");
        Cell titleCell1 = titleRow.createCell(1);
        titleCell1.setCellValue("姓名");
        Cell titleCell2 = titleRow.createCell(2);
        titleCell2.setCellValue("出生日期");
        //循环添加内容
        Row contentRow = null;
        Cell contentCell = null;
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy-MM-dd");
        for(int i=0;i<users.size();i++){
            contentRow = sheet.createRow(i+1);
            contentCell = contentRow.createCell(0);
            contentCell.setCellValue(users.get(i).getUserId());
            contentCell = contentRow.createCell(1);
            contentCell.setCellValue(users.get(i).getUserName());
            contentCell = contentRow.createCell(2);
            contentCell.setCellValue(simpleDateFormat.format(users.get(i).getUserId()));
        }
        OutputStream out = response.getOutputStream();
        response.setContentType("application/ms-excel;charset=UTF-8");
        response.setHeader("Content-Disposition", "attachment;filename="
                .concat(String.valueOf(URLEncoder.encode( fileName, "UTF-8"))));
        work.write(out);


    }

    /**
     * 获取单元格的值，并转换为string类型
     * @param cell
     * @return
     */
    private String getCellValueToString(Cell cell) {
        String cellValue = "";
        if(cell == null){
            return cellValue;
        }
        CellType cellType = cell.getCellTypeEnum();
        //把数字当成String来读，避免出现1读成1.0的情况
        if(cellType.equals("NUMERIC")){
            cell.setCellType(CellType.STRING);
        }
        switch (cellType){
            case _NONE://空
                break;
            case BLANK://空(可能包含空字符)
                break;
            case ERROR://错误
                cellValue = "非法字符";
                break;
            case STRING://字符串
                cellValue = String.valueOf(cell.getStringCellValue());
                break;
            case BOOLEAN://Boolean
                cellValue = String.valueOf(cell.getBooleanCellValue());
                break;
            case FORMULA://公式
                cellValue = String.valueOf(cell.getCellFormula());
                break;
            case NUMERIC://数字
                cellValue = String.valueOf(cell.getNumericCellValue());
                break;
            default:
                cellValue = "未知类型";
                break;
        }
        return cellValue;
    }

    /**
     * 从MultipartFile建立工作簿对象
     * @param file
     * @return
     */
    private Workbook getWorkBook(MultipartFile file) throws IOException {
        //获得文件名
        String fileName = file.getOriginalFilename();
        //创建Workbook工作薄对象，表示整个excel
        Workbook workbook = null;
        //获取excel文件的io流
        InputStream is = file.getInputStream();
        //根据文件后缀名不同(xls和xlsx)获得不同的Workbook实现类对象
        if(fileName.endsWith("xls")){
            //2003
            workbook = new HSSFWorkbook(is);
        }else if(fileName.endsWith("xlsx")){
            //2010
            workbook = new XSSFWorkbook(is);
        }
        return workbook;
    }

    /**
     * 对文件参数进行校验
     * @param file
     * @return
     */
    private int checkFile(MultipartFile file) {
        //判断文件是否存在
        if(null == file){
            return -1;//未获得上传文件！
        }
        //获得文件名
        String fileName = file.getOriginalFilename();
        //判断文件是否是excel文件
        if(!fileName.endsWith("xls") && !fileName.endsWith("xlsx")){
            return -2;//不是有效的excel文件
        }
        return 0;
    }

    class User{
        private int userId;
        private String userName;
        private Date birthDate;

        public int getUserId() {
            return userId;
        }

        public void setUserId(int userId) {
            this.userId = userId;
        }

        public String getUserName() {
            return userName;
        }

        public void setUserName(String userName) {
            this.userName = userName;
        }

        public Date getBirthDate() {
            return birthDate;
        }

        public void setBirthDate(Date birthDate) {
            this.birthDate = birthDate;
        }
    }
}
