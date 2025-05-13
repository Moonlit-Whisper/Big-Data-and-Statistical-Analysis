package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * ClassName: DataComparator
 * Package: org.example
 * Description:
 *
 * @Author 不白之鸢
 * @Create 2025/5/13 22:35
 * @Version 1.0
 */
public class DataComparator {
    public static void DataCompare() {

        // 调用 ExcelDataImport 方法获取上下限数据
        //下线索引是0,上线索引是1
        double[][] doubleArray = ExcelDataImporter.ExcelDataImport();

        // 定义 Excel 文件的路径（相对于项目根目录）
        String filePath = "resources/The original data for samples #285 and #313”.xlsx";

        // 使用 try-with-resources 自动关闭资源
        try (FileInputStream fis = new FileInputStream(filePath); // 打开文件输入流
             Workbook workbook = new XSSFWorkbook(fis)) { // 创建 Excel 工作簿对象

            // 获取 Excel 的第一个工作表（Sheet）
            Sheet sheet = workbook.getSheetAt(4);

            // 获取工作表的最后一行索引
            int lastRowNum = sheet.getLastRowNum();

            // 遍历工作表的每一行，从第1行开始（跳过标题行）
            for (int i = 3; i <= lastRowNum; i++) {
                // 获取当前行
                Row row = sheet.getRow(i);

                // 如果当前行为空，则跳过
                if (row == null) continue;

                // 获取第n+1列的单元格（列）
                Cell dataCell = row.getCell(1);

                // 如果任一单元格为空，则跳过
                if (dataCell == null) continue;

                // 读取第n+1列的内容（浮点数）
                double data = dataCell.getNumericCellValue();
\
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

/*
                // 输出结果
                System.out.println("位号: " + data); // 位号
                if (lowerBound != null && upperBound != null) {
                    // 如果上下限解析成功，输出上下限
                    System.out.println("下限: " + lowerBound + ", 上限: " + upperBound);
                } else {
                    // 如果解析失败，输出错误信息
                    System.out.println("范围解析失败：" + range);
                }
                System.out.println("-------------------------"); // 分隔符
*/
            }

        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
