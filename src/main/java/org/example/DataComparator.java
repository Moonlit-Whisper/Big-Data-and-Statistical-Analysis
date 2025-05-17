package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;

/**
 * ClassName: DataComparator
 * Package: org.example
 * Description:数据比较。
 *              记录下所有不符合范围的数据
 *              以及所在位置（行列），
 *              并且把问题单元格标红。
 *
 * @Author 不白之鸢
 * @Create 2025/5/13 22:35
 * @Version 3.0
 */
public class DataComparator {
    public static void DataCompare() {

        //调用 sigma 方法获取sigma和均值
        //前354为平均值，后2*254为sigma
        BigDecimal[][][] sigma = StatsUtils.sigma();

        // 调用 ExcelDataImport 方法获取上下限数据
        //下限索引是0,上限索引是1
        double[][] doubleArray = ExcelDataImporter.ExcelDataImport();

        // 定义一个计数器，用于统计有问题的单元格数量
        int problematicCellCount = 0;
        // 定义一个计数器，用于统计所有单元格数量
        int totalCellCount = 0;


        // 定义 Excel 输入文件的路径（相对于项目根目录）
        String filePath = "resources/The original data for samples #285 and #313”.xlsx";
        // 定义 Excel 输出文件的路径（相对于项目根目录）
        String outputFilePath = "resources/modified_data.xlsx";


        // 使用 try-with-resources 自动关闭资源
        try (FileInputStream fis = new FileInputStream(filePath); // 打开文件输入流
             FileOutputStream fos = new FileOutputStream(outputFilePath); // 打开文件输出流
             Workbook workbook = new XSSFWorkbook(fis)) { // 创建 Excel 工作簿对象

            // 创建红色单元格样式
            CellStyle redCellStyle = createRedCellStyle(workbook);

            // 获取 Excel 的第一个工作表（Sheet）
            Sheet sheet = workbook.getSheetAt(4);

            // 获取工作表的最后一行索引
            int lastRowNum = sheet.getLastRowNum();

            // 遍历工作表的每一行，从Excel中的第4行开始（跳过标题行）
            // 这i+1是Excel中的行数
            for (int i = 3; i <= lastRowNum; i++) {

                int k = (i < 44) ? 0 : 1;

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

                        // 问题统计自增
                        problematicCellCount++;

                        System.out.println("第" + (i + 1) + "行," + "第" + (j + 1) + "列的单元格数据:" + "空值" + "有问题");

                        continue;
                    }

                    // 读取第j+1列的内容（浮点数）
                    double data = dataCell.getNumericCellValue();

                    if (data < doubleArray[0][j-1] || data > doubleArray[1][j-1] || bessel(data,sigma[0][k][j-1],sigma[1][k][j-1])) {

                        // 问题统计自增
                        problematicCellCount++;

                        // 应用红色样式
                        dataCell.setCellStyle(redCellStyle);

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
            System.out.println();

            // 写入修改后的工作簿到文件
            workbook.write(fos);

            System.out.println("涂色标记完成，结果已保存到:" + outputFilePath);
            System.out.println("问题单元格涂色标记为红");

        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    // 定义一个方法，用于创建红色单元格样式
    private static CellStyle createRedCellStyle(Workbook workbook) {
        CellStyle redCellStyle = workbook.createCellStyle();
        redCellStyle.setFillForegroundColor(IndexedColors.RED.getIndex());
        redCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return redCellStyle;
    }

    private static boolean bessel (double data,BigDecimal average, BigDecimal sigma) {

        // 转换第j+1列的内容（字符串）
        String dataSt = String.valueOf(data);

        // 字符串类型转换为精确小数BigDecimal
        BigDecimal xi = new BigDecimal(dataSt);

        BigDecimal difference = xi.subtract(average);

        BigDecimal absDifference = difference.abs();

        BigDecimal multiple = new BigDecimal("3");

        BigDecimal sigma3 = sigma.multiply(multiple);

        // 比较 absDifference 是否大于 3 sigma
        if (absDifference.compareTo(sigma3) > 0) {
            return true;//absDifference 大于 3 sigma
        } else {
            return false;//absDifference 小于等于 3 sigma
        }
    }
}
