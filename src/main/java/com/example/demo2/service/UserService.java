package com.example.demo2.service;

import com.example.demo2.dao.UserDao;
import com.example.demo2.entity.User;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * 作者: shiloh
 * 日期: 2019/11/8
 * 描述: service
 */
@Service
public class UserService {

    @Autowired
    private UserDao dao;

    private final static String XLS = "xls";

    /**
     * 批量导入
     *
     * @param myFile
     * @throws Exception
     */
    @SuppressWarnings("resource")
    public void batchSaveEquipment(HttpServletRequest request, HttpServletResponse response, MultipartFile myFile) throws Exception {
        List<User> list = new ArrayList<>();
        //获得文件名
        Workbook workbook = null;
        String fileName = myFile.getOriginalFilename();
        System.out.println("文件名：" + fileName);
        if (fileName.endsWith(XLS)) {
            workbook = WorkbookFactory.create(myFile.getInputStream());
//            new XSSFWorkbook(myFile.getInputStream());
        } else {
            throw new Exception("文件格式错误");
        }
        //获取导入的工作表
        Sheet sheet = workbook.getSheet("Sheet1");
        int rows = sheet.getLastRowNum();//获取一共有多少行
        if (rows == 0) {
            throw new Exception("空表格");
        }
        for (int i = 1; i < rows + 1; i++) {
            //读取左上端单元格
            Row row = sheet.getRow(i);
            //避免空列加入或空字符串加入，设置条件
            if (row != null && !"".equals(getCellValue(row.getCell(0)))) {
                //设置成对象
                User user = new User();
                //一行的每格
                user.setXuhao(getCellValue(row.getCell(0)));
                user.setName(getCellValue(row.getCell(1)));
                user.setMobile(getCellValue(row.getCell(2)));
                user.setJiedao(getCellValue(row.getCell(3)));
                user.setJuweihui(getCellValue(row.getCell(4)));
                user.setXiaoqu(getCellValue(row.getCell(5)));
                user.setLouhao(getCellValue(row.getCell(6)));
                user.setDanyuan(getCellValue(row.getCell(7)));
                user.setMenpai(getCellValue(row.getCell(8)));
                user.setHuma(getCellValue(row.getCell(9)));
                user.setJimihuma(getCellValue(row.getCell(10)));
                //加进集合
                list.add(user);
            }
        }
        System.out.println("数据量——" + list.size());
        //设置分批插入
        //暂存数据集合
        List<User> list1 = new ArrayList<>();
        int itemNum = 1000;//每次插入数量
        int numData = 0;//当前数据下标
        //根据总的数据量计算需要分多少次事务插入；
        int numTrans = list.size() / itemNum + 1;
        //  外层循环，j代表提交事务次序；
        for (int j = 1; j <= numTrans; j++) {
            // 从索引numData开始查找总数为itemNum个数据，即为本批要插入的数据量；
            for (int i = numData; i < numData + itemNum; i++) {
                if (i == list.size()) {
                    break;
                }
                list1.add(list.get(i));
            }
            //设置当前数据下标
            numData += itemNum;
            dao.batchInsert(list1);
            //清空暂存数据的集合
            list1.clear();
        }
        exportExcel1(list);

    }

    /**
     * 导出表格
     *
     * @param list
     * @throws IOException
     */
    public void exportExcel1(List<User> list) throws Exception {
        //  文件夹名
        String xiaoqu = list.get(0).getXiaoqu();
        //在D盘创建文件夹，excel导出的文件位置
        File file = new File("D:\\xslExcel\\" + xiaoqu);
        if (!file.exists()) {//判断文件夹是否存在
            file.mkdirs();//创建文件夹
        }
        //输出流
        FileOutputStream os = null;
        //获得所有楼号
        Set<String> sizeSet = new HashSet<>();
        list.stream().forEach(user -> {
            sizeSet.add(user.getLouhao());
        });
        List<String> houseList = new ArrayList<>(sizeSet);
        //循环导出excel到临时文件夹中
        List<User> users = new ArrayList<>();
        for (String number : houseList) {
            //获得单个楼的用户集合
            for (User user : list) {
                if (number.equals(user.getLouhao())) {
                    users.add(user);
                }
            }

//            for (int i = 0; i < list.size(); i++) {
//                if (list.get(i).getLouhao() == number && number.equals(list.get(i).getLouhao())) {
//                    users.add(list.get(i));
//                }
//            }
            //创建webbook
            HSSFWorkbook wb = new HSSFWorkbook();
            //添加sheet,对应表格中的工作表
            HSSFSheet sheet = wb.createSheet("Sheet1");
            //添加第0行(就是表头)
            HSSFRow row = sheet.createRow(0);
/*            //创建单元格设置居中
            HSSFCellStyle style = wb.createCellStyle();
            style.setAlignment(HSSFCellStyle.ALIGN_CENTER);*/
            //设置第一行内容并设置格式为string
            HSSFCell cell = row.createCell(0);
            cell.setCellValue("住户姓名");
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);

            cell = row.createCell(1);
            cell.setCellValue("联系方式");
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);

            cell = row.createCell(2);
            cell.setCellValue("小区名称");
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);

            cell = row.createCell(3);
            cell.setCellValue("楼号");
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);

            cell = row.createCell(4);
            cell.setCellValue("单元号");
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);

            cell = row.createCell(5);
            cell.setCellValue("门牌号");
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);

            cell = row.createCell(6);
            cell.setCellValue("户码");
            cell.setCellType(HSSFCell.CELL_TYPE_STRING);

            //读取数据并写入
            for (int i = 0; i < users.size(); i++) {
                row = sheet.createRow(i + 1);
                User user = users.get(i);
                //创建单元格，并设置值
                row.createCell(0).setCellValue(user.getName());

                row.createCell(1).setCellValue(user.getMobile());
                row.createCell(2).setCellValue(user.getXiaoqu());
                row.createCell(3).setCellValue(Double.valueOf(user.getLouhao()).intValue());
                row.createCell(4).setCellValue(Double.valueOf(user.getDanyuan()).intValue());
                row.createCell(5).setCellValue(Double.valueOf(user.getMenpai()).intValue());
                row.createCell(6).setCellValue(user.getHuma());
            }
            // 导出excel的全路径
            String fullFilePath = file.getAbsolutePath() + File.separator + xiaoqu + (Double.valueOf(number).intValue()) + ".xls";
            //输出excel文件
            os = new FileOutputStream(fullFilePath);
            wb.write(os);
            os.flush();
            //关闭流，清空集合
            users.clear();
            os.close();
            wb.close();
        }
//        output.close();
    }

    /**
     * 导出表格
     *
     * @param request
     * @param response
     * @param list
     * @throws IOException
     */
    public void exportExcel(HttpServletRequest request, HttpServletResponse response, List<User> list) throws Exception {
        // 文件名称
        String filename = list.get(0).getXiaoqu() + list.get(0).getLouhao();
        //创建webbook
        HSSFWorkbook wb = new HSSFWorkbook();
        //添加sheet,对应表格中的工作表
        HSSFSheet sheet = wb.createSheet("Sheet1");
        //添加第0行(就是表头)
        HSSFRow row = sheet.createRow(0);
        //创建单元格设置居中
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);
        //设置第一行内容
        HSSFCell cell = row.createCell(0);
        cell.setCellValue("住户姓名");
        cell.setCellStyle(style);
        cell.setCellValue("联系方式");
        cell.setCellStyle(style);
        cell.setCellValue("小区名称");
        cell.setCellStyle(style);
        cell.setCellValue("楼号");
        cell.setCellStyle(style);
        cell.setCellValue("单元号");
        cell.setCellStyle(style);
        cell.setCellValue("门牌号");
        cell.setCellStyle(style);
        cell.setCellValue("户码");
        cell.setCellStyle(style);
        //读取数据并写入
        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(i + 1);
            User user = list.get(0);
            //创建单元格，并设置值
            row.createCell(0).setCellValue(user.getName());
            row.createCell(1).setCellValue(user.getMobile());
            row.createCell(2).setCellValue(user.getXiaoqu());
            row.createCell(3).setCellValue(user.getLouhao());
            row.createCell(4).setCellValue(user.getDanyuan());
            row.createCell(5).setCellValue(user.getMenpai());
            row.createCell(6).setCellValue(user.getHuma());
        }
        //输出excel文件
        OutputStream output = response.getOutputStream();
        response.reset();
        response.setHeader("Content-disposition", "attachment;filename=" + filename + ".xls");
        response.setContentType("application/msexcel");
        wb.write(output);
        output.close();
    }


    /**
     * 获得cell内容
     *
     * @param cell
     * @return
     * @throws Exception
     */
    public String getCellValue(Cell cell) throws Exception {
        String value = "";
        if (cell != null) {
            switch (cell.getCellType()) {
                case Cell.CELL_TYPE_NUMERIC:
                    value = cell.getNumericCellValue() + "";
                    if (HSSFDateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue();
                        if (date != null) {
                            value = new SimpleDateFormat("yyyy-MM-dd").format(date);
                        } else {
                            value = "";
                        }
                    } else {
                        value = new DecimalFormat("0").format(cell.getNumericCellValue());
                    }
                    break;
                case Cell.CELL_TYPE_STRING:
                    value = cell.getStringCellValue();
                    break;
                case Cell.CELL_TYPE_BOOLEAN:
                    value = cell.getBooleanCellValue() + "";
                    break;
                case Cell.CELL_TYPE_FORMULA:
                    value = cell.getCellFormula() + "";
                    break;
                case Cell.CELL_TYPE_BLANK:
                    value = "";
                    break;
                case Cell.CELL_TYPE_ERROR:
                    value = "非法字符";
                    break;
                default:
                    value = "未知类型";
                    break;
            }
        }
        return value.trim();
    }


}
