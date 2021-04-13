package com.jackxu.excelToSql;


import cn.hutool.core.util.StrUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author jackxu
 */
public class excelToSql {


    public static void main(String[] args) {
        String originAddress = "D:/特权配置(生产) 0413.xlsx";
        String tableName = "tb_woxie_rights_config";
        String generateAddress = "D:/生成的sql.txt";
        int columnCount = 7;
        int type = 1;
        if (type == 1) {
            exportExcel(originAddress, tableName, generateAddress, columnCount);
        } else if (type == 2) {
            exportConsole(originAddress, columnCount, tableName);
        }

    }


    public static void exportExcel(String originAddress, String tableName, String generateAddress, int columnCount) {
        File excelFile = new File(originAddress);
        List<String> sqlList = getSqlList(excelFile, columnCount, tableName);
        OutputStream outPutStream;
        File sqlFile = new File(generateAddress);
        try {
            if (!sqlFile.exists()) {
                sqlFile.createNewFile();
            }
            outPutStream = new FileOutputStream(sqlFile);
            StringBuilder stringBuilder = new StringBuilder();
            for (String sql : sqlList) {
                stringBuilder.append(sql).append("\n");
            }
            // 将可变字符串变为固定长度的字符串，方便下面的转码
            String context = stringBuilder.toString();
            // 因为中文可能会乱码，这里使用了转码，转成UTF-8
            byte[] bytes = context.getBytes("UTF-8");
            // 开始写入内容到文件
            outPutStream.write(bytes);
            // 一定要关闭输出流
            outPutStream.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }


    public static void exportConsole(String originAddress, int columnCount, String tableName) {
        File file = new File(originAddress);
        List<String> sqlList = getSqlList(file, columnCount, tableName);
        for (String sql : sqlList) {
            System.out.println(sql);
        }
    }


    /**
     * @param file        文件
     * @param columnCount 列的数量
     * @param tableName   表名称
     * @return
     */
    public static List<String> getSqlList(File file, int columnCount, String tableName) {
        List<String> sqlList = new ArrayList<>();
        Workbook book = null;
        try {
            InputStream is = new FileInputStream(file.getAbsolutePath());
            book = getWorkbook(is, file.getName());
            // 取第一个sheet
            Sheet sheet = book.getSheetAt(0);
            if (sheet == null) {
                return sqlList;
            }
            int rowCount = sheet.getLastRowNum();

            Row column = sheet.getRow(0);
            List<String> nameList = new ArrayList<>();
            for (int i = 0; i < columnCount; i++) {
                nameList.add(column.getCell(i).toString());
            }

            for (int r = 1; r <= rowCount; r++) {
                Row row = sheet.getRow(r);
                if (row == null) {
                    continue;
                }
                List<Object> valueList = new ArrayList<>();
                for (int i = 0; i < columnCount; i++) {
                    if (row.getCell(i) == null || StrUtil.isBlank((row.getCell(i).toString()))) {
                        valueList.add(null);
                    } else if (CellType.NUMERIC == row.getCell(i).getCellTypeEnum()) {
                        valueList.add(row.getCell(i).getNumericCellValue());
                    } else if (CellType.STRING == row.getCell(i).getCellTypeEnum()) {
                        valueList.add(row.getCell(i).getStringCellValue());
                    } else {
                        valueList.add("不能识别类型");
                    }
                }

                StringBuilder sql = new StringBuilder();
                sql.append("INSERT INTO ");
                sql.append(tableName);
                sql.append(" (");
                for (int i = 0; i < nameList.size(); i++) {
                    sql.append(nameList.get(i));
                    if (i < nameList.size() - 1) {
                        sql.append(", ");
                    }
                }
                sql.append(") VALUES (");
                for (int i = 0; i < valueList.size(); i++) {
                    if (valueList.get(i) == null) {
                        sql.append("\'\'");
                    }
                    if (valueList.get(i) instanceof String) {
                        sql.append("\'");
                        sql.append(valueList.get(i));
                        sql.append("\'");
                    }
                    if (valueList.get(i) instanceof Double) {
                        Double valDouble = Double.valueOf(valueList.get(i).toString());
                        if (isIntegerForDouble(valDouble)) {
                            sql.append(valDouble.intValue());
                        } else {
                            sql.append(valueList.get(i));
                        }
                    }
                    if (i < valueList.size() - 1) {
                        sql.append(", ");
                    }
                }
                sql.append(");");
                sqlList.add(sql.toString());
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return sqlList;
    }


    public static boolean isIntegerForDouble(double obj) {
        // 精度范围
        double eps = 1e-10;
        return obj - Math.floor(obj) < eps;
    }


    public static Workbook getWorkbook(InputStream inStr, String fileName) throws IOException {
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        if (".xls".equals(fileType)) {
            wb = new HSSFWorkbook(inStr); // 2003-
        } else if (".xlsx".equals(fileType)) {
            wb = new XSSFWorkbook(inStr); // 2007+
        }
        return wb;
    }

}
