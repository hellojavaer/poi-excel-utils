package org.hellojavaer.poi.excel.utils.read;

import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.serializer.SerializerFeature;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hellojavaer.poi.excel.utils.*;

import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ReadDemo2 {

    public static void main(String[] args) throws FileNotFoundException {
        InputStream in = ReadDemo2.class.getResourceAsStream("/excel/xlsx/data_file1.xlsx");
        ExcelReadSheetProcessor<TestBean> sheetProcessor = new ExcelReadSheetProcessor<TestBean>() {

            @Override
            public void beforeProcess(ExcelReadContext<TestBean> context) {

            }

            @Override
            public void process(ExcelReadContext<TestBean> context, List<TestBean> list) {
                System.out.println(JSONObject.toJSONString(list, SerializerFeature.WriteDateUseDateFormat));
            }

            @Override
            public void onException(ExcelReadContext<TestBean> context, ExcelReadException e) {
                throw e;
            }

            @Override
            public void afterProcess(ExcelReadContext<TestBean> context) {

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
                 throw new ExcelReadException("test throw exception");
                //return cellValue.getStringValue() + "=>row:" + context.getCurRowIndex() + ",col："
                //       + context.getCurColStrIndex();
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
        sheetProcessor.setTargetClass(TestBean.class);// required
        sheetProcessor.setFieldMapping(fieldMapping);// required
        sheetProcessor.setPageSize(2);//
        sheetProcessor.setTrimSpace(true);
        sheetProcessor.setHeadRowIndex(0);
        sheetProcessor.setRowProcessor(new ExcelReadRowProcessor<TestBean>() {

            public TestBean process(ExcelProcessController controller, ExcelReadContext<TestBean> context, Row row,
                                    TestBean t) {
                return t;
            }
        });

        try {
            ExcelUtils.read(in, sheetProcessor);
        } catch (ExcelReadException e) {
            System.out.println(ExcelReadExceptionFormat.format(e));
        }
    }
}
