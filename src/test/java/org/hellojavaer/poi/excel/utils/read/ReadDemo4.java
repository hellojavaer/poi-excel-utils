package org.hellojavaer.poi.excel.utils.read;

import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.serializer.SerializerFeature;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hellojavaer.poi.excel.utils.ExcelProcessController;
import org.hellojavaer.poi.excel.utils.ExcelUtils;
import org.hellojavaer.poi.excel.utils.TestEnum;

import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ReadDemo4 {

    public static void main(String[] args) throws FileNotFoundException {
        InputStream in = ReadDemo4.class.getResourceAsStream("/excel/xlsx/data_file1.xlsx");
        ExcelReadSheetProcessor<HashMap> sheetProcessor = new ExcelReadSheetProcessor<HashMap>() {

            @Override
            public void beforeProcess(ExcelReadContext<HashMap> context) {

            }

            @Override
            public void process(ExcelReadContext<HashMap> context, List<HashMap> list) {
                System.out.println(JSONObject.toJSONString(list, SerializerFeature.WriteDateUseDateFormat));
            }

            @Override
            public void onException(ExcelReadContext<HashMap> context, ExcelReadException e) {
                throw e;
            }

            @Override
            public void afterProcess(ExcelReadContext<HashMap> context) {

            }
        };
        ExcelReadFieldMapping fieldMapping = new ExcelReadFieldMapping();
        fieldMapping.put("byte", "byteField").setRequired(true);
        fieldMapping.put("short", "shortField");
        fieldMapping.put("int", "intField");
        fieldMapping.put("long", "longField");
        fieldMapping.put("float", "floatField");
        fieldMapping.put("double", "doubleField");
        fieldMapping.put("boolean", "boolField");
        fieldMapping.put("string", "stringField");
        fieldMapping.put("date", "dateField");

        fieldMapping.put("enum1", "enumField1").setCellProcessor(new ExcelReadCellProcessor() {

            public Object process(ExcelReadContext<?> context, Cell cell, ExcelCellValue cellValue) {
                // throw new ExcelReadException("test throw exception");
                return cellValue.getStringValue() + "=>row:" + context.getCurRowIndex() + ",col："
                       + context.getCurColStrIndex();
            }
        });

        ExcelReadCellValueMapping valueMapping = new ExcelReadCellValueMapping();
        valueMapping.put("Please select", null);
        valueMapping.put("Option1", TestEnum.AA.toString());
        valueMapping.put("Option2", TestEnum.BB.toString());
        valueMapping.put("Option3", TestEnum.CC.toString());
        // valueMapping.setDefaultValueWithDefaultInput();
        fieldMapping.put("enum2", "enumField2").setValueMapping(valueMapping).setRequired(false);

        sheetProcessor.setSheetIndex(0);// required.it can be replaced with 'setSheetName(sheetName)';
        sheetProcessor.setStartRowIndex(1);//
        // sheetProcessor.setRowEndIndex(3);//
        sheetProcessor.setTargetClass(HashMap.class);// required
        sheetProcessor.setFieldMapping(fieldMapping);// required
        sheetProcessor.setPageSize(2);//
        sheetProcessor.setTrimSpace(true);
        sheetProcessor.setHeadRowIndex(0);
        sheetProcessor.setRowProcessor(new ExcelReadRowProcessor<HashMap>() {

            @Override
            public HashMap process(ExcelProcessController controller, ExcelReadContext<HashMap> context, Row row,
                                   HashMap t) {
                return t;
            }
        });

        ExcelUtils.read(in, sheetProcessor);
    }
}
