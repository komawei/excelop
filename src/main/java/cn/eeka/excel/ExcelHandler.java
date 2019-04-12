package cn.eeka.excel;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.StringUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @ClassName: ExcelHandler
 * @Description: TODO
 * @author: Dan&Dan
 * @date: 2019/4/12 21:47
 */
public class ExcelHandler {

    private static final int COL_NUM = 5;

    private static final int START_INDEX = 11;

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map<String, Integer> relationTimeMap = new HashMap<>();
        relationTimeMap.put("父", 0);
        relationTimeMap.put("母", 1);
        relationTimeMap.put("妹", 2);
        relationTimeMap.put("夫", 3);
        relationTimeMap.put("妻", 4);
        relationTimeMap.put("长子", 5);
        relationTimeMap.put("次子", 6);
        relationTimeMap.put("媳", 7);
        relationTimeMap.put("婿", 8);
        relationTimeMap.put("长女", 9);
        relationTimeMap.put("次女", 10);
        relationTimeMap.put("长孙", 11);
        relationTimeMap.put("次孙", 12);
        relationTimeMap.put("小孙", 13);
        relationTimeMap.put("曾孙", 14);

        FileInputStream is = new FileInputStream("e:\\test.xlsx");

        Workbook workbook = new XSSFWorkbook(is);
        Sheet sheet = workbook.getSheetAt(0);

        int lastRowNum = sheet.getLastRowNum();

        List<Integer> nonHostRowNumList = new ArrayList<>();

        int lastHostRowNum = 0;

        for (int i = 1; i < lastRowNum; i++) {
            System.out.println("lineNumber->" + i);
            Row row = sheet.getRow(i);
            Cell certNoCell = row.getCell(0);
            if (certNoCell == null) {
                continue;
            }

            certNoCell.setCellType(CellType.STRING);

            String certNo = certNoCell.getStringCellValue();
            if (StringUtils.isEmpty(certNo)) {
                continue;
            }

            Cell relationCell = row.getCell(7);
            if (relationCell == null) {
                continue;
            }
            String relation = relationCell.getStringCellValue();

            // 户主跳过，记下当前户主的行数
            if ("户主".equals(relation)) {
                lastHostRowNum = i;
                continue;
            }

            nonHostRowNumList.add(i);

            Integer crossTimes = relationTimeMap.get(relation);
            if (crossTimes == null) {
                continue;
            }

            Cell hostNameCell = row.getCell(1);
            String hostName = hostNameCell.getStringCellValue();

            Cell idCardCell = row.getCell(2);
            String idCard = idCardCell.getStringCellValue();

            int stockHolderNameCellIndex = START_INDEX + crossTimes * COL_NUM;
            int relationHostRelationCellIndex = stockHolderNameCellIndex + 1;
            int relationIdCardNoCellIndex = relationHostRelationCellIndex + 1;
            int relationStockAmountCellIndex = relationIdCardNoCellIndex + 1;

            // 设置户主所在行相应的亲属信息
            Row hostRow = sheet.getRow(lastHostRowNum);
            Cell stockHolderNameCell = hostRow.getCell(stockHolderNameCellIndex);
            Cell relationHostRelationCell = hostRow.getCell(relationHostRelationCellIndex);
            Cell relationIdCardNoCell = hostRow.getCell(relationIdCardNoCellIndex);
            Cell relationStockAmountCell = hostRow.getCell(relationStockAmountCellIndex);

            stockHolderNameCell.setCellValue(hostName);
            relationHostRelationCell.setCellValue(relation);
            relationIdCardNoCell.setCellValue(idCard);

        }

        // 删除非户主所在行
        for (int j = nonHostRowNumList.size() - 1; j < 0; j--) {
//            Row nonHostRow = sheet.getRow(j);
            sheet.shiftColumns(j, j, -1);
        }

        FileOutputStream excelFileOutPutStream = new FileOutputStream("E:\\test.xlsx");

        // 将最新的 Excel 文件写入到文件输出流中，更新文件信息！

        workbook.write(excelFileOutPutStream);

        // 执行 flush 操作， 将缓存区内的信息更新到文件上

        excelFileOutPutStream.flush();

        // 使用后，及时关闭这个输出流对象， 好习惯，再强调一遍！
        is.close();
        excelFileOutPutStream.close();

    }
}
