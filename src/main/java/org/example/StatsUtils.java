package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.math.RoundingMode;

/**
 * ClassName: StatsUtils
 * Package: org.example
 * Description: 统计工具类。
 *             计算均值和sigma。
 *            传回的数据最高维是分均值与sigma
 *            最高维索引为0的表示均值。
 *            第二维分两个对象，
 *            第三维分每个单元格。
 *
 *
 * @Author 不白之鸢
 * @Create 2025/5/15 15:05
 * @Version 2.0
 */
public class StatsUtils {
    public static BigDecimal[][][] sigma() {

        // 定义一个 BigDecimal 数组,用于存储,第一个是平均值,第二个是sigma，
        BigDecimal[][][] bigDecimalArrayReturn = new BigDecimal[2][2][354];

        BigDecimal nBig = new BigDecimal("40");
        BigDecimal nSmall = new BigDecimal("39");
        int scale = 30; // 精确到小数点后30位


        // 定义 Excel 输入文件的路径（相对于项目根目录）
        String filePath = "resources/The original data for samples #285 and #313”.xlsx";

        BigDecimal[][] rootArray = new BigDecimal[2][354];

        BigDecimal[] bigDecimalSumArray = new BigDecimal[354];

        BigDecimal[] bigDecimalProductSumArray = new BigDecimal[354];

        BigDecimal[][] bigDecimalSumAverageArray = new BigDecimal[2][354];

        for (int i = 0; i < 354; i++) {

            bigDecimalSumArray[i] = new BigDecimal(0);

            bigDecimalProductSumArray[i] = new BigDecimal(0);

        }

        // 使用 try-with-resources 自动关闭资源
        try (FileInputStream fis = new FileInputStream(filePath); // 打开文件输入流
             Workbook workbook = new XSSFWorkbook(fis)) { // 创建 Excel 工作簿对象

            // 获取 Excel 的第一个工作表（Sheet）
            Sheet sheet = workbook.getSheetAt(4);

            // 获取工作表的最后一行索引
            int lastRowNum = sheet.getLastRowNum();//82

            for (int m = 0; m < 2; m++) {

                int lowerBound, upperBound;
                int mid = (lastRowNum - 3 >> 1) + 3;//42

                switch (m) {

                    case 0:
                        lowerBound = 3;
                        upperBound = mid;
                        break;

                    case 1:
                        lowerBound = mid + 1;
                        upperBound = lastRowNum;
                        //清洗
                        clear(bigDecimalProductSumArray);
                        clear(bigDecimalSumArray);
                        break;

                    default:
                        throw new IllegalStateException("Unexpected value: " + m);
                }

                // 遍历工作表的每一行，从Excel中的第4行开始（跳过标题行）
                // 这i+1是Excel中的行数
                for (int i = lowerBound; i <= upperBound; i++) {
                    // 获取当前行
                    Row row = sheet.getRow(i);

                    // 如果当前行为空，则跳过
                    if (row == null) continue;

                    // 这j+1是Excel中的列数
                    for (int j = 1; j <= 354; j++) {

                        // 从Excel中的第2列开始
                        // 获取第j+1列的单元格（列）
                        Cell dataCell = row.getCell(j);

                        // 如果任一单元格为空，则报空值并跳过
                        if (dataCell == null) continue;

                        // 读取第j+1列的内容（浮点数）
                        double data = dataCell.getNumericCellValue();

                        // 转换第j+1列的内容（字符串）
                        String dataSt = String.valueOf(data);

                        // 字符串类型转换为精确小数BigDecimal
                        BigDecimal tem = new BigDecimal(dataSt);

                        // 计算Xi平方
                        BigDecimal product = tem.multiply(tem);

                        // 计算所有Xi平方的和
                        bigDecimalProductSumArray[j - 1] = bigDecimalProductSumArray[j - 1].add(product);

                        // 计算所有Xi(每列)的和
                        bigDecimalSumArray[j - 1] = bigDecimalSumArray[j - 1].add(tem);

                    }
                }

                for (int h = 0; h < 354; h++) {

                    // 计算每列和的平方
                    BigDecimal bigDecimalSumProduct = bigDecimalSumArray[h].multiply(bigDecimalSumArray[h]);

                    BigDecimal division1 = bigDecimalSumProduct.divide(nBig, scale, RoundingMode.HALF_UP);

                    BigDecimal difference = bigDecimalProductSumArray[h].subtract(division1);

                    BigDecimal division2 = difference.divide(nSmall, scale, RoundingMode.HALF_UP);

                    rootArray[m][h] = sqrt(division2, scale);

                    bigDecimalSumAverageArray[m][h] = bigDecimalSumArray[h].divide(nBig, scale, RoundingMode.HALF_UP);

                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }

        for (int j = 0; j < 2; j++) {
            for (int h = 0; h < 354; h++) {

                bigDecimalArrayReturn[0][j][h] = bigDecimalSumAverageArray[j][h]; // 平均值
            }

        }

        for (int j = 0; j < 2; j++) {
            for (int h = 0; h < 354; h++) {

                bigDecimalArrayReturn[1][j][h] = rootArray[j][h]; // sigma
            }

        }


        return bigDecimalArrayReturn;
    }

    public static BigDecimal sqrt(BigDecimal value, int scale) {
        if (value.compareTo(BigDecimal.ZERO) < 0) {
            throw new ArithmeticException("Cannot calculate square root of a negative value");
        }
        if (value.compareTo(BigDecimal.ZERO) == 0) {
            return BigDecimal.ZERO;
        }

        BigDecimal two = BigDecimal.valueOf(2);
        BigDecimal guess = value.divide(two, scale, RoundingMode.HALF_UP);
        if (guess.compareTo(BigDecimal.ZERO) == 0) {
            // value 很小导致 guess 变成 0，直接用 1 作为初始 guess
            guess = BigDecimal.ONE;
        }
        BigDecimal epsilon = BigDecimal.valueOf(1).scaleByPowerOfTen(-scale);

        while (true) {
            BigDecimal nextGuess = guess.add(value.divide(guess, scale, RoundingMode.HALF_UP))
                    .divide(two, scale, RoundingMode.HALF_UP);
            if (guess.subtract(nextGuess).abs().compareTo(epsilon) <= 0) {
                break;
            }
            guess = nextGuess;
            if (guess.compareTo(BigDecimal.ZERO) == 0) {
                // 防止迭代过程中 guess 变为 0
                throw new ArithmeticException("Guess became zero during sqrt computation");
            }
        }
        return guess;
    }

    //清洗数组
    public static void clear(BigDecimal[] a) {
        for (int i = 0; i < a.length; i++) {
            a[i] = new BigDecimal(0);
        }
    }
}
