package com.fengy.upload.service.impl;

import com.fengy.upload.entity.ExcelEntity;
import com.fengy.upload.service.ExcelService;
import com.fengy.upload.util.ExcelUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.util.StringUtils;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.*;

@Service
public class ExcelServiceImpl implements ExcelService {

    private final static Integer COLNUM = 7;

    private final static Map<String,String> COLMAP = new LinkedHashMap<>();

    {
        COLMAP.put("银行名","0");
        COLMAP.put("代码","1");
        COLMAP.put("合计","2");
        COLMAP.put("同业","3");
        COLMAP.put("贴现","4");
        COLMAP.put("券","5");
        COLMAP.put("金融券","6");
    }

    @Override
    public Map<String, Object> getExcelData(MultipartFile file) {
        Map<String,Object> resMap = new LinkedHashMap<>();
        //获取文件名称
        String fileName = file.getOriginalFilename();
        System.out.println(fileName);

        try {
            //获取输入流
            InputStream in = file.getInputStream();
            //判断excel版本
            Workbook workbook = null;
            if (ExcelUtil.judegExcelEdition(fileName)) {
                workbook = new XSSFWorkbook(in);
            } else {
                workbook = new HSSFWorkbook(in);
            }

            //遍历所有sheet
            for (int z = 0; z < workbook.getNumberOfSheets(); z++) {
                //获取sheet对象
                Sheet sheet = workbook.getSheetAt(z);
                //创建当前sheet信息集合
                Map<String, ExcelEntity> result = new LinkedHashMap<>();

                Row row=null;
                for (int i=1; i<sheet.getPhysicalNumberOfRows();i++) {

                    row = sheet.getRow(i);  //row从第二行开始获取
                    if (row == null) continue;

                    //循环获取每一列
                    ArrayList<String> list = new ArrayList<>();
                    Cell cell = null;
                    for (int j = 0; j < COLNUM; j++) {
                        cell = row.getCell(j);
                        if(cell==null){
                            cell = row.createCell(j);
                            cell.setCellValue("");
                        }
                        cell.setCellType(CellType.STRING);

                        Object value = cell.getStringCellValue();

                        if (COLMAP.containsKey(value.toString())) continue;  //如果是标题


                        list.add(cell.getStringCellValue());
                    }
                    ExcelEntity entity = listToEntity(list);
                    resMap = putResData(resMap,entity);

                }

                //将装有每一列的集合装入大集合
                //关闭资源
                workbook.close();
            }
        } catch (FileNotFoundException e) {
            e.printStackTrace();
            System.out.println("===================未找到文件======================");
        } catch (IOException e) {
            e.printStackTrace();
            System.out.println("===================上传失败======================");
        }
        return resMap;
    }

    @Override
    public String exportExcel(String olename, Map<String, Object> resMap, HttpServletResponse response) {

        Workbook hwb  = null;
        if (ExcelUtil.judegExcelEdition(olename)) {
            hwb  = new XSSFWorkbook();
        } else {
            hwb  = new HSSFWorkbook();
        }
        //声明一个单子并命名
        Sheet sheet = hwb.createSheet("汇总数据");
        //给单子名称一个长度
        sheet.setDefaultColumnWidth((short)15);
        //生成一个样式
        CellStyle style = hwb.createCellStyle();
        //创建第一行（也可以成为表头）
        Row row = sheet.createRow(0);
        Cell col_cell = null;

        for (Map.Entry<String, String> col : COLMAP.entrySet()) {
            Object col_name = col.getKey();
            Object col_num = col.getValue();
            col_cell = row.createCell(Integer.parseInt(col_num.toString()));
            col_cell.setCellValue(col_name.toString());
            col_cell.setCellStyle(style);

        }
        int rowcount = 1;

        Double sum_tongye = new Double("0");
        Double sum_tiexian = new Double("0");
        Double sum_juan = new Double("0");
        Double sum_jinrongjuan = new Double("0");

        for (Map.Entry<String, Object> entry : resMap.entrySet()) {
            ExcelEntity entity = (ExcelEntity) entry.getValue();
            row =sheet.createRow(rowcount++);
            row.createCell(0).setCellValue(entity.getName());
            row.createCell(1).setCellValue(entity.getCode());
            row.createCell(2).setCellValue(entity.getHeji());
            row.createCell(3).setCellValue(entity.getTongye());
            row.createCell(4).setCellValue(entity.getTiexian());
            row.createCell(5).setCellValue(entity.getJuan());
            row.createCell(6).setCellValue(entity.getJinrongjuan());

            sum_tongye = add(sum_tongye,entity.getTongye());
            sum_tiexian = add(sum_tiexian,entity.getTiexian());
            sum_juan = add(sum_juan,entity.getJuan());
            sum_jinrongjuan = add(sum_jinrongjuan,entity.getJinrongjuan());
        }

        row = sheet.createRow(rowcount);
        row.createCell(3).setCellValue(sum_tongye);
        row.createCell(4).setCellValue(sum_tiexian);
        row.createCell(5).setCellValue(sum_juan);
        row.createCell(6).setCellValue(sum_jinrongjuan);

        try {
            hwb.write(response.getOutputStream());
            hwb.close();
        } catch (IOException e) {
            e.printStackTrace();
        }


        return "";
    }

    private Map<String, Object> putResData(Map<String, Object> resMap, ExcelEntity entity) {
        if (StringUtils.isEmpty(entity.getCode())) return resMap;

        String code = entity.getCode().toUpperCase().trim();
        if (resMap.containsKey(code)){
            ExcelEntity oldEntity = (ExcelEntity) resMap.get(code);
            Double heji = add(oldEntity.getHeji(),entity.getHeji());
            Double tongye = add(oldEntity.getTongye(),entity.getTongye());
            Double tiexian = add(oldEntity.getTiexian(),entity.getTiexian());
            Double juan = add(oldEntity.getJuan(),entity.getJuan());
            Double jinrongjuan = add(oldEntity.getJinrongjuan(),entity.getJinrongjuan());

            oldEntity.setHeji(heji);
            oldEntity.setTongye(tongye);
            oldEntity.setTiexian(tiexian);
            oldEntity.setJuan(juan);
            oldEntity.setJinrongjuan(jinrongjuan);

        }else{
            resMap.put(code,entity);
        }
        return resMap;
    }

    private ExcelEntity listToEntity(ArrayList<String> list) {
        ExcelEntity entity = new ExcelEntity();
        Double sum_value = new Double("0");
        for (int i = 0; i < list.size(); i++) {
            if (i > 1){
                Double value;
                try {
                    value =  Double.valueOf(list.get(i));
                }catch (Exception e){
                    value = new Double("0");
                }
                sum_value = add(sum_value,value);
            }
            switch (i){
                case 0:
                    entity.setName(list.get(0));
                    break;
                case 1:
                    entity.setCode(list.get(1).toUpperCase());
                    break;
                case 3:
                    entity.setTongye(changeStrToNum(list.get(3)));
                    break;
                case 4:
                    entity.setTiexian(changeStrToNum(list.get(4)));
                    break;
                case 5:
                    entity.setJuan(changeStrToNum(list.get(5)));
                    break;
                case 6:
                    entity.setJinrongjuan(changeStrToNum(list.get(6)));
                    break;

            }
        }
        entity.setHeji(sum_value);
        return entity;
    }

    private Double changeStrToNum(String s) {
        Double value = new Double("0");
        if (StringUtils.isEmpty(s)){
            return value;
        }
        try {
            value = Double.valueOf(s);
        }catch (NumberFormatException e){
            value = new Double("0");
        }
        return value;
    }

    /**
     * 格式化，double保留两位小数
     * @param price
     * @return
     */
    public static String formatDouble(double price){
        BigDecimal b = new BigDecimal(price);
        return  b.setScale(2, BigDecimal.ROUND_HALF_UP) + "";
    }
    /**
     * * 两个Double数相加 *
     *
     * @param v1 *
     * @param v2 *
     * @return Double
     */
    public static Double add(Double v1, Double v2) {
        if (StringUtils.isEmpty(v1)) v1 = new Double("0");
        if (StringUtils.isEmpty(v2)) v2 = new Double("0");

        BigDecimal b1 = new BigDecimal(v1.toString());
        BigDecimal b2 = new BigDecimal(v2.toString());
        return new Double(b1.add(b2).doubleValue());
    }
    /**
     * * 四个Double数相加 *
     *
     * @param v1 *
     * @param v2 *
     * @param v3 *
     * @param v4 *
     * @return Double
     */
    public static Double getSum(Double v1, Double v2,Double v3, Double v4) {
        if (StringUtils.isEmpty(v1)) v1 = new Double("0");
        if (StringUtils.isEmpty(v2)) v2 = new Double("0");
        if (StringUtils.isEmpty(v3)) v2 = new Double("0");
        if (StringUtils.isEmpty(v4)) v2 = new Double("0");

        BigDecimal b1 = new BigDecimal(v1.toString());
        BigDecimal b2 = new BigDecimal(v2.toString());
        BigDecimal b3 = new BigDecimal(v3.toString());
        BigDecimal b4 = new BigDecimal(v4.toString());
        return new Double(b1.add(b2).add(b3).add(b4).doubleValue());
    }



}
