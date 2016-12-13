package org.hellojavaer.poi.excel.utils.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hellojavaer.poi.excel.utils.ExcelProcessController;
import org.hellojavaer.poi.excel.utils.ExcelUtils;
import org.hellojavaer.poi.excel.utils.TestBean;
import org.hellojavaer.poi.excel.utils.TestEnum;
import org.hellojavaer.poi.excel.utils.read.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class WriteDemo2 {

    private static List<TestBean> testDataCache;

    public static void main(String[] args) throws IOException {
        InputStream excelTemplate = WriteDemo2.class.getResourceAsStream("/excel/xlsx/template_file3.xlsx");
        URL url = WriteDemo2.class.getResource("/");
        final String outputFilePath = url.getPath() + "output_file2.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.createNewFile();
        FileOutputStream output = new FileOutputStream(outputFile);

        final AtomicLong rowIndex = new AtomicLong(0);
        ExcelWriteSheetProcessor<TestBean> sheetProcessor = new ExcelWriteSheetProcessor<TestBean>() {

            @Override
            public void beforeProcess(ExcelWriteContext<TestBean> context) {
                System.out.println("write excel start!");
            }

            @Override
            public void onException(ExcelWriteContext<TestBean> context, ExcelWriteException e) {
                throw e;
            }

            @Override
            public void afterProcess(ExcelWriteContext<TestBean> context) {
                context.setCellValue(2, 0, "Test output");
                context.setCellValue(4, 0, "zoukaiming");
                context.setCellValue(6, 0, "hellojavaer@gmail.com");
                context.setCellValue(8, 0, new Date());
                System.out.println("write excel end!");
                System.out.println("output file path is " + outputFilePath);
            }
        };
        ExcelWriteFieldMapping fieldMapping = new ExcelWriteFieldMapping();
        fieldMapping.put("C", "byteField");
        fieldMapping.put("D", "shortField");
        fieldMapping.put("E", "intField");
        fieldMapping.put("F", "longField");
        fieldMapping.put("G", "floatField");
        fieldMapping.put("H", "doubleField");
        fieldMapping.put("I", "boolField");
        fieldMapping.put("J", "stringField");
        fieldMapping.put("K", "dateField");
        fieldMapping.put("L", "enumField1").setCellProcessor(new ExcelWriteCellProcessor<TestBean>() {

            public void process(ExcelWriteContext<TestBean> context, Object obj, Cell cell) {
                if (obj == null) {
                    cell.setCellValue("Please select");
                }
            }
        });
        ExcelWriteCellValueMapping kValueMapping = new ExcelWriteCellValueMapping();
        kValueMapping.put(null, "Please select");
        kValueMapping.put(TestEnum.AA.toString(), "Option1");
        kValueMapping.put(TestEnum.BB.toString(), "Option2");
        kValueMapping.put(TestEnum.CC.toString(), "Option3");
        fieldMapping.put("M", "enumField2").setValueMapping(kValueMapping);

        sheetProcessor.setSheetIndex(0);
        sheetProcessor.setStartRowIndex(1);
        sheetProcessor.setFieldMapping(fieldMapping);
        sheetProcessor.setTemplateRows(1, 2);
        // sheetProcessor.setRowProcessor(new ExcelWriteRowProcessor<TestBean>() {
        // @Override
        // public void process(ExcelProcessController controller, ExcelWriteContext<TestBean> context, TestBean t,
        // Row row) {
        // }
        // });
        sheetProcessor.setDataList(getDataList());

        ExcelUtils.write(excelTemplate, output, sheetProcessor);
    }

    private static List<TestBean> getDataList() {
        final List<TestBean> re = new ArrayList<TestBean>();
        InputStream in = WriteDemo2.class.getResourceAsStream("/excel/xlsx/data_file2.xlsx");
        ExcelReadSheetProcessor<TestBean> sheetProcessor = new ExcelReadSheetProcessor<TestBean>() {

            @Override
            public void beforeProcess(ExcelReadContext<TestBean> context) {

            }

            @Override
            public void process(ExcelReadContext<TestBean> context, List<TestBean> list) {
                re.addAll(list);
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
        fieldMapping.put("A", "byteField");
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
                String str = cellValue.getStringValue();
                if (str == null || str.trim().equals("")) {
                    return null;
                } else {
                    return str;
                }
            }
        });

        ExcelReadCellValueMapping valueMapping = new ExcelReadCellValueMapping();
        valueMapping.put("Please select", null);
        valueMapping.put("Option1", TestEnum.AA.toString());
        valueMapping.put("Option2", TestEnum.BB.toString());
        valueMapping.put("Option3", TestEnum.CC.toString());
        fieldMapping.put("K", "enumField2").setValueMapping(valueMapping).setRequired(false);

        sheetProcessor.setSheetIndex(0);// required.it can be replaced with 'setSheetName(sheetName)';
        sheetProcessor.setStartRowIndex(1);//
        sheetProcessor.setTargetClass(TestBean.class);// required
        sheetProcessor.setFieldMapping(fieldMapping);// required
        sheetProcessor.setRowProcessor(new ExcelReadRowProcessor<TestBean>() {

            public TestBean process(ExcelProcessController controller, ExcelReadContext<TestBean> context, Row row,
                                    TestBean t) {
                return t;
            }
        });

        ExcelUtils.read(in, sheetProcessor);
        return re;
    }
}
