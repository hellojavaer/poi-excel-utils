package org.hellojavaer.poi.excel.utils.read;

import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hellojavaer.poi.excel.utils.ExcelProcessController;
import org.hellojavaer.poi.excel.utils.ExcelUtils;
import org.hellojavaer.poi.excel.utils.TestBean;
import org.hellojavaer.poi.excel.utils.TestEnum;

import com.alibaba.fastjson.JSONObject;
import com.alibaba.fastjson.serializer.SerializerFeature;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ReadDemo1 {

    public static void main(String[] args) throws FileNotFoundException {
        InputStream in = ReadDemo1.class.getResourceAsStream("/excel/xlsx/data_file1.xlsx");
        ExcelReadSheetProcessor<TestBean> sheetProcessor = new ExcelReadSheetProcessor<TestBean>() {

            @Override
            public void beforeProcess(ExcelReadContext<TestBean> context) {

            }

            @Override
            public void process(ExcelReadContext<TestBean> context, List<TestBean> list) {
                System.out.println(JSONObject.toJSONString(list, SerializerFeature.WriteDateUseDateFormat));
            }

            @Override
            public void onExcepton(ExcelReadContext<TestBean> context, RuntimeException e) {
                if (e instanceof ExcelReadException) {
                    ExcelReadException ere = (ExcelReadException) e;
                    if (ere.getCode() == ExcelReadException.CODE_OF_CELL_VALUE_REQUIRED) {
                        System.out.println("at row:" + (ere.getRowIndex() + 1) + " column:" + ere.getColStrIndex()
                                           + ", data cant't be null.");
                    } else if (ere.getCode() == ExcelReadException.CODE_OF_CELL_VALUE_NOT_MATCHED) {
                        System.out.println("at row:" + (ere.getRowIndex() + 1) + " column:" + ere.getColStrIndex()
                                           + ", data doesn't match.");
                    } else if (ere.getCode() == ExcelReadException.CODE_OF_CELL_ERROR) {
                        System.out.println("at row:" + (ere.getRowIndex() + 1) + " column:" + ere.getColStrIndex()
                                           + ", cell error.");
                    } else {
                        System.out.println("at row:" + (ere.getRowIndex() + 1) + " column:" + ere.getColStrIndex()
                                           + ", process error. detail message is: " + ere.getMessage());
                    }
                } else {
                    throw e;
                }
            }

            @Override
            public void afterProcess(ExcelReadContext<TestBean> context) {

            }
        };
        ExcelReadFieldMapping fieldMapping = new ExcelReadFieldMapping();
        fieldMapping.put("A", "byteField");
        fieldMapping.put("B", "shortField");
        fieldMapping.put("C", "intField");
        fieldMapping.put("D", "longField");
        fieldMapping.put("E", "floatField");
        fieldMapping.put("F", "doubleField");
        fieldMapping.put("G", "boolField");
        fieldMapping.put("H", "stringField");
        fieldMapping.put("I", "dateField");

        fieldMapping.put("J", "enumField1", new ExcelReadCellProcessor() {

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
        fieldMapping.put("K", "enumField2", valueMapping, false);

        sheetProcessor.setSheetIndex(0);// required.it can be replaced with 'setSheetName(sheetName)';
        sheetProcessor.setRowStartIndex(1);//
        // sheetProcessor.setRowEndIndex(3);//
        sheetProcessor.setTargetClass(TestBean.class);// required
        sheetProcessor.setFieldMapping(fieldMapping);// required
        sheetProcessor.setPageSize(2);//
        sheetProcessor.setSkipEmptyRow(true);
        sheetProcessor.setTrimSpace(true);
        sheetProcessor.setRowProcessor(new ExcelReadRowProcessor<TestBean>() {

            // if return null, null will not be added to reasult data list.
            public TestBean process(ExcelProcessController controller, ExcelReadContext<TestBean> context, Row row,
                                    TestBean t) {
                return t;
            }
        });

        ExcelUtils.read(in, sheetProcessor);
    }
}
