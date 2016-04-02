package org.hellojavaer.poi.excel.utils.read;

import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.hellojavaer.poi.excel.utils.ExcelProcessController;
import org.hellojavaer.poi.excel.utils.ExcelUtils;
import org.hellojavaer.poi.excel.utils.TestBean;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ReadDemo3 {

    public static void main(String[] args) throws FileNotFoundException {
        InputStream in = ReadDemo3.class.getResourceAsStream("/excel/xlsx/data_file1.xlsx");
        ExcelReadSheetProcessor<TestBean> sheetProcessor = new ExcelReadSheetProcessor<TestBean>() {

            @Override
            public void beforeProcess(ExcelReadContext<TestBean> context) {
            }

            @Override
            public void process(ExcelReadContext<TestBean> context, List<TestBean> list) {
            }

            @Override
            public void onExcepton(ExcelReadContext<TestBean> context, RuntimeException e) {
                throw e;
            }

            @Override
            public void afterProcess(ExcelReadContext<TestBean> context) {
            }
        };

        sheetProcessor.setSheetIndex(0);// required.it can be replaced with 'setSheetName(sheetName)';
        sheetProcessor.setStartRowIndex(1);//
        // sheetProcessor.setRowEndIndex(3);//
        sheetProcessor.setTargetClass(TestBean.class);// required
        sheetProcessor.setPageSize(2);//
        sheetProcessor.setRowProcessor(new ExcelReadRowProcessor<TestBean>() {

            public TestBean process(ExcelProcessController controller, ExcelReadContext<TestBean> context, Row row,
                                    TestBean t) {
                if (row != null) {
                    Cell cell = row.getCell(0);
                    System.out.println(ExcelUtils.readCell(cell));
                }
                return t;
            }
        });
        ExcelUtils.read(in, sheetProcessor);
    }
}
