package com.jackxu.excelToSql;


import cn.hutool.core.util.StrUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
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
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class excelToSql {

    public static void main(String[] args) {
        File excelFile = new File("D:/特权配置(生产) 0412 去掉钻石无忧取消版本.xlsx");
        String tableName = "tb_home_page_rights";
        List<String> sqlList = getSqlList(excelFile, 7, tableName);
        OutputStream outPutStream;
        File sqlFile = new File("D:/生成的sql.txt");
        try {
            // 处理文件已存在的情况
            if (sqlFile.exists()) {
                return;
            }
            sqlFile.createNewFile();
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


    public static void homePageToConsole() {
        File file = new File("D:/excelToDB0407.xlsx");
        List<String> sqlList = getSqlList(file, 7, "tb_home_page_rights");
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
                List<String> valueList = new ArrayList<>();
                for (int i = 0; i < columnCount; i++) {
                    if (row.getCell(i) == null || StrUtil.isBlank((row.getCell(i).toString()))) {
                        valueList.add("");
                    } else if (isContainChinese(row.getCell(i).toString()) || row.getCell(i).toString().contains("*")) {
                        valueList.add(row.getCell(i).toString());
                    } else {
                        // stripTrailingZeros()去除尾部多余的0
                        // toPlainString()不使用任何指数
                        valueList.add(new BigDecimal(row.getCell(i).toString()).stripTrailingZeros().toPlainString());
                    }
//					if (row.getCell(i) == null) {
//						valueList.add("");
//					} else if (row.getCell(i).toString().contains(".") || isNumeric(row.getCell(i).toString())) {
//						// stripTrailingZeros()去除尾部多余的0
//						// toPlainString()不使用任何指数
//						valueList.add(new BigDecimal(row.getCell(i).toString()).stripTrailingZeros().toPlainString());
//					} else {
//						valueList.add(row.getCell(i).toString());
//					}
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
                    if (StrUtil.isNotBlank(valueList.get(i))) {
                        sql.append("\'");
                        sql.append(valueList.get(i));
                        sql.append("\'");
                    } else {
                        sql.append("\'\'");
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

    /**
     * 是否包含中文
     *
     * @param str
     * @return
     */
    public static boolean isContainChinese(String str) {
        Pattern p = Pattern.compile("[\u4e00-\u9fa5]");
        Matcher m = p.matcher(str);
        if (m.find()) {
            return true;
        }
        return false;
    }

    public static boolean isNumeric(String str) {
        Pattern pattern = Pattern.compile("[0-9]*");
        Matcher isNum = pattern.matcher(str);
        if (!isNum.matches()) {
            return false;
        }
        return true;
    }
}
