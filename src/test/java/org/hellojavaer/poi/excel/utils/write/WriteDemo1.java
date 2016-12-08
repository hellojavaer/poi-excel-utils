package org.hellojavaer.poi.excel.utils.write;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.hellojavaer.poi.excel.utils.ExcelType;
import org.hellojavaer.poi.excel.utils.ExcelUtils;
import org.hellojavaer.poi.excel.utils.TestBean;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.concurrent.atomic.AtomicBoolean;

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
                System.out.println("write excel start!");
            }

            @Override
            public void onException(ExcelWriteContext<TestBean> context, ExcelWriteException e) {
                throw e;
            }

            @Override
            public void afterProcess(ExcelWriteContext<TestBean> context) {

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

    private static List<TestBean> getDateList() {
        List<TestBean> list = new ArrayList<TestBean>();
        TestBean bean = new TestBean();
        bean.setByteField((byte) 1);
        bean.setShortField((short) 2);
        bean.setIntField(3);
        bean.setLongField(4L);
        bean.setFloatField(5.1f);
        bean.setDoubleField(6.23);
        bean.setBoolField(true);
        bean.setEnumField1("enumField1");
        bean.setEnumField2("enumField2");
        bean.setDateField(new Date());
        bean.setStringField("test");

        list.add(bean);
        list.add(bean);
        list.add(bean);
        return list;
    }
}
