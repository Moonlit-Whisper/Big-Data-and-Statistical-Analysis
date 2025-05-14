package org.example;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;

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

        // 定义一个计数器，用于统计有问题的单元格数量
        int problematicCellCount = 0;
        // 定义一个计数器，用于统计所有单元格数量
        int totalCellCount = 0;


        // 调用 ExcelDataImport 方法获取上下限数据
        //下限索引是0,上限索引是1
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

            // 遍历工作表的每一行，从Excel中的第4行开始（跳过标题行）
            // 这i+1是Excel中的行数
            for (int i = 3; i <= lastRowNum; i++) {
                // 获取当前行
                Row row = sheet.getRow(i);

                // 如果当前行为空，则跳过
                if (row == null) continue;

                // 这j+1是Excel中的列数
                for (int j = 1; j <= 354; j++) {

                    // 总数统计自增
                    totalCellCount++;

                    // 从Excel中的第2列开始
                    // 获取第j+1列的单元格（列）
                    Cell dataCell = row.getCell(j);

                    // 如果任一单元格为空，则报空值并跳过
                    if (dataCell == null) {
                        problematicCellCount++;// 问题统计自增
                        System.out.println("第" + (i + 1) + "行," + "第" + (j + 1) + "列的单元格数据:" + "空值" + "有问题");
                        continue;
                    }

                    // 读取第j+1列的内容（浮点数）
                    double data = dataCell.getNumericCellValue();

                    if (data < doubleArray[0][j-1] || data > doubleArray[1][j-1]) {
                        problematicCellCount++;// 问题统计自增
                        // 如果数据不在上下限范围内，则输出提示信息
                        System.out.println("第" + (i + 1) + "行," + "第" + (j + 1) + "列的单元格数据:" + data + "有问题");
                        System.out.println("-------------------------"); // 分隔符
                    }
                }
            }

            // 输出检查结果
            System.out.println();
            System.out.println("应检查" + 80 * 354 + "个数据单元格。");
            System.out.println("实检查" + totalCellCount + "个数据单元格。");
            System.out.println("其中，" + problematicCellCount + "个单元格数据有问题" + "，可查看上面的提示信息" + "。");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }
}
