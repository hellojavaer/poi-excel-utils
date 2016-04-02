package org.hellojavaer.poi.excel.utils.write;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.hellojavaer.poi.excel.utils.ExcelType;
import org.hellojavaer.poi.excel.utils.ExcelUtils;
import org.hellojavaer.poi.excel.utils.TestBean;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class WriteDemo1 {

    public static void main(String[] args) throws IOException {
        URL url = WriteDemo1.class.getResource("/");
        final String outputFilePath = url.getPath() + "output_file1.xlsx";
        File outputFile = new File(outputFilePath);
        outputFile.createNewFile();
        FileOutputStream output = new FileOutputStream(outputFile);
        final AtomicBoolean b = new AtomicBoolean(false);

        ExcelWriteSheetProcessor<TestBean> sheetProcessor = new ExcelWriteSheetProcessor<TestBean>() {

            @Override
            public void beforeProcess(ExcelWriteContext<TestBean> context) {
                for (int col = 0; col <= 8; col++) {
                    context.setCellValue(0, col, "field" + col);
                }
                System.out.println("write excel start!");
            }

            @Override
            public List<TestBean> getDataList(ExcelWriteContext<TestBean> context) {
                if (b.get() == false) {
                    b.set(true);
                    List<TestBean> list = new ArrayList<TestBean>();
                    TestBean bean = new TestBean();
                    bean.setStringField("test");
                    bean.setLongField(11L);
                    bean.setBoolField(true);
                    bean.setByteField((byte) 1);
                    bean.setDateField(new Date());
                    bean.setDoubleField(1.23);
                    bean.setEnumField1("enumField1");
                    bean.setEnumField2("enumField2");
                    bean.setFloatField(2345678901.1f);
                    bean.setIntField(1111);
                    bean.setLongField(2222L);
                    bean.setShortField((short) 33);
                    bean.setStringField("stringField");

                    list.add(bean);
                    list.add(bean);
                    list.add(bean);
                    return list;
                } else {
                    return null;
                }
            }

            @Override
            public void onException(ExcelWriteContext<TestBean> context, RuntimeException e) {
                if (e instanceof ExcelWriteException) {
                    ExcelWriteException ewe = (ExcelWriteException) e;
                    if (ewe.getCode() == ExcelWriteException.CODE_OF_FIELD_VALUE_NOT_MATCHED) {
                        System.out.println("at row:" + (ewe.getRowIndex() + 1) + " column:" + ewe.getColStrIndex()
                                           + ", data doesn't match.");
                    } else {
                        System.out.println("at row:" + (ewe.getRowIndex() + 1) + " column:" + ewe.getColStrIndex()
                                           + ", process error. detail message is: " + ewe.getMessage());
                    }
                } else {
                    throw e;
                }
            }

            @Override
            public void afterProcess(ExcelWriteContext<TestBean> context) {

                Cell cell = context.setCellValue(context.getCurRowIndex() + 2,
                                                 context.getCurRowIndex() + 4,
                                                 0,
                                                 8,
                                                 "Thinks for using pio-excel-utils,if you have any questions or suggestions when you are useing,"
                                                         + " please connect me.my email is hellojavaer@gmail.com.I'm zoukaimiing.");
                CellStyle cellStyle = cell.getSheet().getWorkbook().createCellStyle();
                cellStyle.setWrapText(true);
                cell.setCellStyle(cellStyle);
                System.out.println("write excel end!");
                System.out.println("output file path is " + outputFilePath);
            }
        };

        ExcelWriteFieldMapping fieldMapping = new ExcelWriteFieldMapping();
        fieldMapping.put("A", "byteField");
        fieldMapping.put("B", "shortField");
        fieldMapping.put("C", "intField");
        fieldMapping.put("D", "longField");
        fieldMapping.put("E", "floatField");
        fieldMapping.put("F", "doubleField");
        fieldMapping.put("G", "boolField");
        fieldMapping.put("H", "stringField");
        fieldMapping.put("I", "dateField");

        sheetProcessor.setSheetIndex(0);// required. It can be replaced with 'setSheetName(sheetName)';
        sheetProcessor.setStartRowIndex(1);//
        sheetProcessor.setFieldMapping(fieldMapping);// required
        // sheetProcessor.setTemplateRowIndex(1);

        ExcelUtils.write(ExcelType.XLSX, output, sheetProcessor);
    }
}
