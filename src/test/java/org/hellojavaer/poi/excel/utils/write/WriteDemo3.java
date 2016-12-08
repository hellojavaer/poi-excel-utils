package org.hellojavaer.poi.excel.utils.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.hellojavaer.poi.excel.utils.ExcelType;
import org.hellojavaer.poi.excel.utils.ExcelUtils;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.*;
import java.util.concurrent.atomic.AtomicBoolean;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class WriteDemo3 {

    public static void main(String[] args) throws IOException {
        URL url = WriteDemo3.class.getResource("/");
        final String outputFilePath = url.getPath() + "output_file3.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.createNewFile();
        FileOutputStream output = new FileOutputStream(outputFile);
        final AtomicBoolean b = new AtomicBoolean(false);

        ExcelWriteSheetProcessor<Map> sheetProcessor = new ExcelWriteSheetProcessor<Map>() {

            @Override
            public void beforeProcess(ExcelWriteContext<Map> context) {
                System.out.println("write excel start!");
            }

            @Override
            public void onException(ExcelWriteContext<Map> context, ExcelWriteException e) {
                throw e;
            }

            @Override
            public void afterProcess(ExcelWriteContext<Map> context) {

                Cell cell = context.setCellValue(context.getCurRowIndex() + 2, context.getCurRowIndex() + 4, 0, 8,
                                                 "Thinks for using pio-excel-utils!");
                CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
                cellStyle.setWrapText(true);
                cell.setCellStyle(cellStyle);
                System.out.println("write excel end!");
                System.out.println("output file path is " + outputFilePath);
            }
        };

        ExcelWriteFieldMapping fieldMapping = new ExcelWriteFieldMapping();
        fieldMapping.put("A", "byteField").setHead("byteField");
        fieldMapping.put("B", "shortField").setHead("shortField");
        fieldMapping.put("C", "intField").setHead("intField");
        fieldMapping.put("D", "longField").setHead("longField");
        fieldMapping.put("E", "floatField").setHead("floatField");
        fieldMapping.put("F", "doubleField").setHead("doubleField");
        fieldMapping.put("G", "boolField").setHead("boolField");
        fieldMapping.put("H", "stringField").setHead("stringField");
        fieldMapping.put("I", "dateField").setHead("dateField");

        sheetProcessor.setSheetIndex(0);// required. It can be replaced with 'setSheetName(sheetName)';
        sheetProcessor.setStartRowIndex(1);//
        sheetProcessor.setFieldMapping(fieldMapping);// required
        sheetProcessor.setHeadRowIndex(0);
        // sheetProcessor.setTemplateRowIndex(1);
        sheetProcessor.setDataList(getDateList());

        ExcelUtils.write(ExcelType.XLSX, output, sheetProcessor);
    }

    private static List<Map> getDateList() {
        List<Map> list = new ArrayList<Map>();
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("byteField", (byte) 1);
        map.put("shortField", (short) 2);
        map.put("intField", 3);
        map.put("longField", 4L);
        map.put("floatField", 5.1f);
        map.put("doubleField", 6.23);
        map.put("boolField", true);
        map.put("dateField", new Date());
        map.put("enumField1", "enumField1");
        map.put("enumField2", "enumField2");
        map.put("stringField", "map_test");
        list.add(map);
        list.add(map);
        list.add(map);
        return list;
    }
}
