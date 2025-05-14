package org.example;

/**
 * ClassName: ExcelDataImporter
 * Package: org.example
 * Description:
 *
 * @Author 不白之鸢
 * @Create 2025/5/7 20:39
 * @Version 2.0
 */


import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelDataImporter {

    public static double[][] ExcelDataImport() {

        // 定义一个二维数组，用于存储 Excel 中的上下限数据，最后返回
        double[][] doubleArray = new double[2][354];

        // 定义 Excel 文件的路径（相对于项目根目录）
        String filePath = "resources/354 operation variable information.xlsx";

        // 使用 try-with-resources 自动关闭资源
        try (FileInputStream fis = new FileInputStream(filePath); // 打开文件输入流
             Workbook workbook = new XSSFWorkbook(fis)) { // 创建 Excel 工作簿对象

            // 获取 Excel 的第一个工作表（Sheet）
            Sheet sheet = workbook.getSheetAt(0);

            // 获取工作表的最后一行索引
            int lastRowNum = sheet.getLastRowNum();

            // 遍历工作表的每一行，从第1行开始（跳过标题行）
            for (int i = 1; i <= lastRowNum; i++) {
                // 获取当前行
                Row row = sheet.getRow(i);

                // 如果当前行为空，则跳过
                if (row == null) continue;

                // 获取第2列的单元格（位号列）
                Cell dataCell = row.getCell(1);

                // 获取第4列的单元格（范围列）
                Cell rangeCell = row.getCell(3);

                // 如果任一单元格为空，则跳过
                if (dataCell == null || rangeCell == null) continue;

                // 读取第2列的内容（位号）
                String data = dataCell.getStringCellValue();

                // 读取第4列的内容（范围字符串，例如 "-1.5-(-1.2)"）
                String range = rangeCell.getStringCellValue();

                // 定义正则表达式，用于匹配范围字符串中的上下限
                Pattern pattern = Pattern.compile("^\\s*(-?\\d+(\\.\\d+)?|[\\(（]-?\\d+(\\.\\d+)?[\\)）])\\s*-\\s*(-?\\d+(\\.\\d+)?|[\\(（]-?\\d+(\\.\\d+)?[\\)）])\\s*$");
                Matcher matcher = pattern.matcher(range);

                // 定义上下限变量
                Double lowerBound = null, upperBound = null;

                // 如果匹配成功，提取上下限
                if (matcher.find()) {
                    lowerBound = parseBound(matcher.group(1)); // 提取下限
                    upperBound = parseBound(matcher.group(4)); // 提取上限
                } else {
                    // 如果范围格式不正确，输出错误信息并跳过
                    System.out.println("范围解析失败：" + range);
                    continue;
                }

                // 如果上下限颠倒，则交换它们
                if (lowerBound != null && upperBound != null && lowerBound > upperBound) {
                    double temp = lowerBound;
                    lowerBound = upperBound;
                    upperBound = temp;
                }

                if (lowerBound != null && upperBound != null) {
                    // 如果上下限解析成功，赋值上下限数组
                    doubleArray[1][i - 1] = upperBound; // 上限
                    doubleArray[0][i - 1] = lowerBound; // 下限
//                    System.out.println("下限: " + doubleArray[0][i - 1] + ", 上限: " +  doubleArray[1][i - 1]);
                } else {
                    // 如果解析失败，输出错误信息
                    System.out.println("范围解析失败：" + range);
                }

            }

        } catch (IOException e) {
            e.printStackTrace();
        }

        return doubleArray; // 返回解析后的上下限数组

    }

    /**
     * 解析范围值，移除括号并转换为 Double 类型
     *
     * @param boundStr 范围字符串
     * @return 转换后的 Double 值
     */
    private static double parseBound(String boundStr) {
        // 移除中文和英文括号
        boundStr = boundStr.replaceAll("[()（）]", "");
        // 将字符串转换为 Double 类型
        return Double.parseDouble(boundStr);
    }
}
