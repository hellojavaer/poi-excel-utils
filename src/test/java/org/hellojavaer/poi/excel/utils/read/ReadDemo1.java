package org.hellojavaer.poi.excel.utils.read;

import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.serializer.SerializerFeature;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hellojavaer.poi.excel.utils.ExcelProcessController;
import org.hellojavaer.poi.excel.utils.ExcelUtils;
import org.hellojavaer.poi.excel.utils.TestBean;
import org.hellojavaer.poi.excel.utils.TestEnum;

import java.io.InputStream;
import java.util.List;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ReadDemo1 {

    public static void main(String[] args) throws Throwable {
        InputStream in = ReadDemo1.class.getResourceAsStream("/excel/xlsx/data_file1.xlsx");
        ExcelReadSheetProcessor<TestBean> sheetProcessor = new ExcelReadSheetProcessor<TestBean>() {

            @Override
            public void beforeProcess(ExcelReadContext<TestBean> context) {
            }

            @Override
            public void process(ExcelReadContext<TestBean> context, List<TestBean> list) {
                System.out.println(JSONObject.toJSONString(list, SerializerFeature.WriteDateUseDateFormat));
                // throw new IllegalStateException("step3.");
                // throw new ExcelReadException("step3.");
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
        fieldMapping.put("A", "byteField").setRequired(true).setLinkField("url");
        fieldMapping.put("B", "shortField");
        fieldMapping.put("C", "intField");
        fieldMapping.put("D", "longField");
        fieldMapping.put("E", "floatField");
        fieldMapping.put("F", "doubleField");
        fieldMapping.put("G", "boolField");
        fieldMapping.put("H", "stringField");
        fieldMapping.put("I", "dateField");

        fieldMapping.put("J", "enumField1").setCellProcessor(new ExcelReadCellProcessor() {

            public Object process(ExcelReadContext<?> context, Cell cell, ExcelCellValue cellValue) {
                // throw new ExcelReadException("step1.");
                // throw new IllegalStateException("step1.");
                return cellValue.getStringValue() + "=>row:" + context.getCurRowIndex() + ",col："
                       + context.getCurColStrIndex();
            }
        });

        ExcelReadCellValueMapping valueMapping = new ExcelReadCellValueMapping();
        valueMapping.put(null, null);
        valueMapping.put("Please select", null);
        valueMapping.put("Option1", TestEnum.AA.toString());
        valueMapping.put("Option2", TestEnum.BB.toString());
        valueMapping.put("Option3", TestEnum.CC.toString());
        // valueMapping.setDefaultValueWithDefaultInput();
        fieldMapping.put("K", "enumField2").setValueMapping(valueMapping);

        sheetProcessor.setSheetIndex(0);// required.it can be replaced with 'setSheetName(sheetName)';
        sheetProcessor.setStartRowIndex(1);//
        // sheetProcessor.setRowEndIndex(3);//
        sheetProcessor.setTargetClass(TestBean.class);// required
        sheetProcessor.setFieldMapping(fieldMapping);// required
        sheetProcessor.setPageSize(2);//
        sheetProcessor.setTrimSpace(true);
        sheetProcessor.setRowProcessor(new ExcelReadRowProcessor<TestBean>() {

            public TestBean process(ExcelProcessController controller, ExcelReadContext<TestBean> context, Row row,
                                    TestBean t) {
                // throw new ExcelReadException("step2.");
                // throw new IllegalStateException("step2.");
                return t;
            }
        });

        try {
            ExcelUtils.read(in, sheetProcessor);
        } catch (ExcelReadException e) {
            // throw e.getCause();
            // System.out.println(e.getMessage());
            e.printStackTrace();
        }
    }
}
