/*
 * Copyright 2015-2016 the original author or authors.
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 *
 *      http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package org.hellojavaer.poi.excel.utils;

import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.Modifier;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Calendar;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;

import org.apache.commons.lang.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hellojavaer.poi.excel.utils.read.ExcelCellValue;
import org.hellojavaer.poi.excel.utils.read.ExcelReadCellValueMapping;
import org.hellojavaer.poi.excel.utils.read.ExcelReadContext;
import org.hellojavaer.poi.excel.utils.read.ExcelReadException;
import org.hellojavaer.poi.excel.utils.read.ExcelReadFieldMapping;
import org.hellojavaer.poi.excel.utils.read.ExcelReadRowProcessor;
import org.hellojavaer.poi.excel.utils.read.ExcelReadSheetProcessor;
import org.hellojavaer.poi.excel.utils.read.InnerReadCellProcessorWrapper;
import org.hellojavaer.poi.excel.utils.write.ExcelWriteCellProcessor;
import org.hellojavaer.poi.excel.utils.write.ExcelWriteCellValueMapping;
import org.hellojavaer.poi.excel.utils.write.ExcelWriteContext;
import org.hellojavaer.poi.excel.utils.write.ExcelWriteException;
import org.hellojavaer.poi.excel.utils.write.ExcelWriteFieldMapping;
import org.hellojavaer.poi.excel.utils.write.ExcelWriteSheetProcessor;
import org.hellojavaer.poi.excel.utils.write.InnerWriteCellProcessorWrapper;
import org.springframework.beans.BeanUtils;
import org.springframework.util.Assert;

import com.alibaba.fastjson.util.TypeUtils;

/**
 * 
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ExcelUtils {

    private static long TIME_1899_12_31_00_00_00_000;
    private static long TIME_1900_01_01_00_00_00_000;
    private static long TIME_1900_01_02_00_00_00_000;

    static {
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss:SSS");
        try {
            TIME_1899_12_31_00_00_00_000 = df.parse("1899-12-31 00:00:00:000").getTime();
            TIME_1900_01_01_00_00_00_000 = df.parse("1900-01-01 00:00:00:000").getTime();
            TIME_1900_01_02_00_00_00_000 = df.parse("1900-01-02 00:00:00:000").getTime();
        } catch (ParseException e) {
            throw new RuntimeException(e);
        }
    }

    private static void readConfigParamVerify(ExcelReadSheetProcessor<?> sheetProcessor) {
        Class<?> clazz = sheetProcessor.getTargetClass();
        ExcelReadFieldMapping fieldMapping = sheetProcessor.getFieldMapping();
        for (Entry<Integer, Map<String, InnerReadCellProcessorWrapper>> indexFieldMapping : fieldMapping.entrySet()) {
            for (Map.Entry<String, InnerReadCellProcessorWrapper> filedMapping : indexFieldMapping.getValue().entrySet()) {
                String fieldName = filedMapping.getKey();
                if (fieldName != null) {
                    PropertyDescriptor pd = BeanUtils.getPropertyDescriptor(clazz, fieldName);
                    if (pd == null || pd.getWriteMethod() == null) {
                        throw new IllegalArgumentException("In fieldMapping config {colIndex:"
                                                           + indexFieldMapping.getKey() + "["
                                                           + convertColIntIndexToCharIndex(indexFieldMapping.getKey())
                                                           + "]<->fieldName:" + filedMapping.getKey() + "}, "
                                                           + " class " + clazz.getName() + " can't find field '"
                                                           + filedMapping.getKey() + "' and can not also find "
                                                           + filedMapping.getKey() + "'s writter method.");
                    }
                    if (!Modifier.isPublic(pd.getWriteMethod().getDeclaringClass().getModifiers())) {
                        pd.getWriteMethod().setAccessible(true);
                    }
                }
            }
        }
    }

    /**
     * parse excel file data to java object
     * 
     * @param workbookInputStream
     * @param sheetProcessors
     */
    @SuppressWarnings({ "unchecked", "rawtypes" })
    public static void read(InputStream workbookInputStream, ExcelReadSheetProcessor<?>... sheetProcessors) {
        Assert.isTrue(workbookInputStream != null, "workbookInputStream can't be null");
        Assert.isTrue(sheetProcessors != null && sheetProcessors.length != 0, "sheetProcessor can't be null");
        try {
            Workbook workbook = WorkbookFactory.create(workbookInputStream);
            for (ExcelReadSheetProcessor<?> sheetProcessor : sheetProcessors) {
                readConfigParamVerify(sheetProcessor);
                ExcelReadContext context = new ExcelReadContext();
                try {
                    Class clazz = sheetProcessor.getTargetClass();
                    Integer sheetIndex = sheetProcessor.getSheetIndex();
                    String sheetName = sheetProcessor.getSheetName();
                    context.setCurSheetIndex(sheetIndex);
                    context.setCurSheetName(sheetName);

                    Sheet sheet = null;
                    if (sheetName != null) {
                        sheet = workbook.getSheet(sheetName);
                        if (sheet != null && sheetIndex != null && !sheetIndex.equals(workbook.getSheetIndex(sheet))) {
                            throw new IllegalArgumentException("sheetName[" + sheetName + "] and sheetIndex["
                                                               + sheetIndex + "] not match.");
                        }
                    } else if (sheetIndex != null) {
                        sheet = workbook.getSheetAt(sheetIndex);
                    } else {
                        throw new IllegalArgumentException("sheetName or sheetIndex can't be null");
                    }
                    if (sheet == null) {
                        throw new IllegalArgumentException("Sheet Not Found Exception. for sheet name:" + sheetName);
                    }

                    if (sheetIndex == null) {
                        sheetIndex = workbook.getSheetIndex(sheet);
                    }
                    if (sheetName == null) {
                        sheetName = sheet.getSheetName();
                    }

                    // proc sheet
                    context.setCurSheet(sheet);
                    context.setCurSheetIndex(sheetIndex);
                    context.setCurSheetName(sheet.getSheetName());
                    context.setCurRow(null);
                    context.setCurRowData(null);
                    context.setCurRowIndex(null);
                    context.setCurColIndex(null);
                    context.setCurColIndex(null);
                    // beforeProcess
                    sheetProcessor.beforeProcess(context);

                    if (sheetProcessor.getPageSize() != null) {
                        context.setDataList(new ArrayList(sheetProcessor.getPageSize()));
                    } else {
                        context.setDataList(new ArrayList());
                    }

                    Integer pageSize = sheetProcessor.getPageSize();
                    int startRow = sheetProcessor.getRowStartIndex();
                    Integer rowEndIndex = sheetProcessor.getRowEndIndex();
                    int actLastRow = sheet.getLastRowNum();
                    if (rowEndIndex != null) {
                        if (rowEndIndex > actLastRow) {
                            rowEndIndex = actLastRow;
                        }
                    } else {
                        rowEndIndex = actLastRow;
                    }
                    if (pageSize != null) {
                        int total = rowEndIndex - startRow + 1;
                        int pageCount = (total + pageSize - 1) / pageSize;
                        for (int i = 0; i < pageCount; i++) {
                            int start = startRow + pageSize * i;
                            int size = pageSize;
                            if (i == pageCount - 1) {
                                size = rowEndIndex - start + 1;
                            }
                            read(context, sheet, start, size, sheetProcessor.getFieldMapping(), clazz,
                                 sheetProcessor.getRowProcessor());
                            sheetProcessor.process(context, context.getDataList());
                            context.getDataList().clear();
                        }
                    } else {
                        read(context, sheet, startRow, rowEndIndex - startRow + 1, sheetProcessor.getFieldMapping(),
                             clazz, sheetProcessor.getRowProcessor());
                        sheetProcessor.process(context, context.getDataList());
                        context.getDataList().clear();
                    }
                } catch (RuntimeException e) {
                    sheetProcessor.onExcepton(context, e);
                } finally {
                    sheetProcessor.afterProcess(context);
                }
            }
        } catch (Exception e) {
            if (e instanceof RuntimeException) {
                throw (RuntimeException) e;
            } else {
                throw new RuntimeException(e);
            }
        }
    }

    private static <T> void read(ExcelReadContext<T> context, Sheet sheet, int startRow, Integer pageSize,
                                 ExcelReadFieldMapping fieldMapping, Class<T> targetClass,
                                 ExcelReadRowProcessor<T> processor) {
        Assert.isTrue(sheet != null, "sheet can't be null");
        Assert.isTrue(startRow >= 0, "startRow must greater than or equal to 0");
        Assert.isTrue(pageSize == null || pageSize >= 1, "pageSize == null || pageSize >= 1");
        Assert.isTrue(fieldMapping != null, "fieldMapping can't be null");
        Assert.isTrue(targetClass != null, "clazz can't be null");

        List<T> list = context.getDataList();
        if (sheet.getPhysicalNumberOfRows() == 0) {
            return;
        }
        //
        int endRow = sheet.getLastRowNum();
        if (pageSize != null) {
            endRow = startRow + pageSize - 1;
        }
        for (int i = startRow; i <= endRow; i++) {
            Row row = sheet.getRow(i);
            // proc row
            context.setCurRow(row);
            context.setCurRowIndex(i);
            context.setCurCell(null);
            context.setCurColIndex(null);

            if (row == null) {
                continue;
            }
            T t = null;
            if (row != null) {
                t = readRow(context, row, fieldMapping, targetClass, processor);
            }
            if (processor != null) {
                try {
                    t = processor.process(context, row, t);
                } catch (RuntimeException re) {
                    if (re instanceof ExcelReadException) {
                        ExcelReadException ere = (ExcelReadException) re;
                        ere.setRowIndex(row.getRowNum());
                        // ere.setColIndex();
                        throw ere;
                    } else {
                        ExcelReadException e = new ExcelReadException(re);
                        e.setRowIndex(row.getRowNum());
                        e.setColIndex(null);
                        e.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                        throw e;
                    }
                }
            }
            if (t != null) {// ignore empty row
                list.add(t);
            }
        }
        // return
    }

    private static <T> T readRow(ExcelReadContext<T> context, Row row, ExcelReadFieldMapping fieldMapping,
                                 Class<T> targetClass, ExcelReadRowProcessor<T> processor) {
        short minColIx = row.getFirstCellNum();
        short maxColIx = row.getLastCellNum();// note ,this return value is
                                              // 1-based.
        short lastColIndex = (short) (maxColIx - 1);
        try {
            context.setCurRowData(targetClass.newInstance());
        } catch (Exception e1) {
            throw new RuntimeException(e1);
        }

        for (Entry<Integer, Map<String, InnerReadCellProcessorWrapper>> fieldMappingEntry : fieldMapping.entrySet()) {
            int curColIndex = fieldMappingEntry.getKey();// excel index;
            // proc cell
            context.setCurColIndex(curColIndex);
            context.setCurCell(null);

            if (curColIndex > lastColIndex || curColIndex < minColIx) {
                Map<String, InnerReadCellProcessorWrapper> fields = fieldMappingEntry.getValue();
                for (Map.Entry<String, InnerReadCellProcessorWrapper> field : fields.entrySet()) {
                    // @SuppressWarnings("unused")
                    // String fieldName = field.getValue().getFieldName();
                    if (field.getValue().isRequired()) {
                        ExcelReadException e = new ExcelReadException();
                        e.setRowIndex(row.getRowNum());
                        e.setColIndex(curColIndex);
                        e.setCode(ExcelReadException.CODE_OF_CELL_VALUE_REQUIRED);
                        throw e;
                    } else {
                        // TODO SET NULL
                    }
                }
            } else {
                Cell cell = row.getCell(curColIndex);
                context.setCurCell(cell);
                Map<String, InnerReadCellProcessorWrapper> fields = fieldMappingEntry.getValue();
                for (Map.Entry<String, InnerReadCellProcessorWrapper> fieldEntry : fields.entrySet()) {
                    String fieldName = fieldEntry.getKey();
                    InnerReadCellProcessorWrapper entry = fieldEntry.getValue();
                    if (cell == null) {
                        if (entry.isRequired()) {
                            ExcelReadException e = new ExcelReadException();
                            e.setRowIndex(row.getRowNum());
                            e.setColIndex(curColIndex);
                            e.setCode(ExcelReadException.CODE_OF_CELL_VALUE_REQUIRED);
                            throw e;
                        } else {
                            continue;
                        }
                    }
                    PropertyDescriptor pd = org.springframework.beans.BeanUtils.getPropertyDescriptor(targetClass,
                                                                                                      fieldName);
                    if (pd == null || pd.getWriteMethod() == null) {
                        continue;
                    }

                    Object value = _readCell(cell);
                    value = procValueConvert(context, row, cell, entry, fieldName, value);
                    if (value != null) {// ignore null
                        try {
                            Class<?> paramType = pd.getWriteMethod().getParameterTypes()[0];
                            if (value != null && !paramType.isAssignableFrom(value.getClass())) {
                                value = TypeUtils.cast(value, paramType, null);
                            }
                            pd.getWriteMethod().invoke(context.getCurRowData(), value);
                        } catch (Exception e1) {
                            ExcelReadException e = new ExcelReadException(e1);
                            e.setRowIndex(row.getRowNum());
                            e.setColIndex(cell.getColumnIndex());
                            e.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                            throw e;
                        }
                    }
                }
            }
        }
        return context.getCurRowData();
    }

    /**
     * convert Cell type to ExcelCellValue type
     * @param cell
     * @return
     * @see ExcelCellValue
     */
    public static ExcelCellValue readCell(Cell cell) {
        if (cell == null) {
            return null;
        } else {
            Object val = _readCell(cell);
            return new ExcelCellValue(val);
        }
    }

    private static Object _readCell(Cell cell) {
        int cellType = cell.getCellType();
        Object value = null;
        switch (cellType) {
            case Cell.CELL_TYPE_BLANK:
                value = null;
                break;
            case Cell.CELL_TYPE_BOOLEAN:
                boolean bool = cell.getBooleanCellValue();
                value = bool;
                break;
            case Cell.CELL_TYPE_ERROR:
                // cell.getErrorCellValue();
                ExcelReadException e = new ExcelReadException();
                e.setRowIndex(cell.getRowIndex());
                e.setColIndex(cell.getColumnIndex());
                e.setCode(ExcelReadException.CODE_OF_CELL_ERROR);
                throw e;
            case Cell.CELL_TYPE_FORMULA:
                String formula = cell.getCellFormula();
                if (StringUtils.isBlank(formula)) {
                    formula = null;
                }
                if (formula != null) {
                    formula = formula.trim();
                }
                value = formula;
                break;
            case Cell.CELL_TYPE_NUMERIC:
                Object inputValue = null;//
                double doubleVal = cell.getNumericCellValue();
                if (DateUtil.isCellDateFormatted(cell)) {
                    inputValue = DateUtil.getJavaDate(doubleVal);
                } else {
                    long longVal = Math.round(cell.getNumericCellValue());
                    if (Double.parseDouble(longVal + ".0") == doubleVal) {
                        inputValue = longVal;
                    } else {
                        inputValue = doubleVal;
                    }
                }
                value = inputValue;
                break;
            case Cell.CELL_TYPE_STRING:
                String str = cell.getStringCellValue();
                if (StringUtils.isBlank(str)) {
                    str = null;
                }
                if (str != null) {
                    str = str.trim();
                }
                value = str;
                break;
            default:
                throw new RuntimeException("unsupport cell type " + cellType);
        }
        return value;
    }

    private static Object procValueConvert(ExcelReadContext<?> context, Row row, Cell cell,
                                           InnerReadCellProcessorWrapper entry, String fieldName, Object value) {
        Object convertedValue = value;
        if (entry.getValueMapping() != null) {
            ExcelReadCellValueMapping valueMapping = entry.getValueMapping();
            String strValue = TypeUtils.castToString(value);
            convertedValue = valueMapping.get(strValue);
            if (convertedValue == null) {
                if (!valueMapping.containsKey(strValue)) {
                    if (valueMapping.isSetDefaultValue()) {
                        if (valueMapping.isSetDefaultValueWithDefaultInput()) {
                            convertedValue = value;
                        } else {
                            convertedValue = valueMapping.getDefaultValue();
                        }
                    } else if (valueMapping.getDefaultProcessor() != null) {
                        try {
                            convertedValue = valueMapping.getDefaultProcessor().process(context, cell,
                                                                                        new ExcelCellValue(value));
                        } catch (RuntimeException re) {
                            if (re instanceof ExcelReadException) {
                                ExcelReadException ere = (ExcelReadException) re;
                                ere.setRowIndex(row.getRowNum());
                                ere.setColIndex(cell.getColumnIndex());
                                throw ere;
                            } else {
                                ExcelReadException e = new ExcelReadException(re);
                                e.setRowIndex(row.getRowNum());
                                e.setColIndex(cell.getColumnIndex());
                                e.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                                throw e;
                            }
                        }
                        if (convertedValue != null && convertedValue instanceof ExcelCellValue) {
                            convertedValue = value;
                        }
                    } else {
                        ExcelReadException e = new ExcelReadException();
                        e.setRowIndex(row.getRowNum());
                        e.setColIndex(cell.getColumnIndex());
                        e.setCode(ExcelReadException.CODE_OF_CELL_VALUE_NOT_MATCHED);
                        throw e;
                    }
                }
            }
        } else if (entry.getProcessor() != null) {
            try {
                convertedValue = entry.getProcessor().process(context, cell, new ExcelCellValue(value));
            } catch (RuntimeException re) {
                if (re instanceof ExcelReadException) {
                    ExcelReadException ere = (ExcelReadException) re;
                    ere.setRowIndex(row.getRowNum());
                    ere.setColIndex(cell.getColumnIndex());
                    throw ere;
                } else {
                    ExcelReadException e = new ExcelReadException(re);
                    e.setRowIndex(row.getRowNum());
                    e.setColIndex(cell.getColumnIndex());
                    e.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                    throw e;
                }
            }
            if (convertedValue != null && convertedValue instanceof ExcelCellValue) {
                convertedValue = value;
            }
        }
        if (convertedValue == null && entry.isRequired()) {
            ExcelReadException e = new ExcelReadException();
            e.setRowIndex(row.getRowNum());
            e.setColIndex(cell.getColumnIndex());
            e.setCode(ExcelReadException.CODE_OF_CELL_VALUE_REQUIRED);
            throw e;
        } else {
            return convertedValue;
        }
    }

    /**
     * parse java object to excel file
     * 
     * @param template
     * @param outputStream
     * @param sheetProcessors
     */
    public static void write(InputStream template, OutputStream outputStream,
                             ExcelWriteSheetProcessor<?>... sheetProcessors) {
        Assert.notNull(template);
        Assert.notNull(outputStream);
        Assert.isTrue(sheetProcessors != null && sheetProcessors.length > 0);
        Workbook workbook;
        try {
            workbook = WorkbookFactory.create(template);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        write(true, workbook, outputStream, sheetProcessors);
    }

    /**
     * parse java object to excel file
     * 
     * @param fileType
     * @param outputStream
     * @param sheetProcessors
     */
    public static void write(ExcelType fileType, OutputStream outputStream,
                             ExcelWriteSheetProcessor<?>... sheetProcessors) {

        Assert.notNull(fileType);
        Assert.notNull(outputStream);
        Assert.isTrue(sheetProcessors != null && sheetProcessors.length > 0);
        Workbook workbook = null;
        if (fileType == ExcelType.XLS) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        write(false, workbook, outputStream, sheetProcessors);
    }

    private static class InnerRow {

        private short                   height;
        private float                   heightInPoints;
        private CellStyle               rowStyle;
        private boolean                 zeroHeight;
        private Map<Integer, InnerCell> cellMap = new HashMap<Integer, InnerCell>();

        public short getHeight() {
            return height;
        }

        public void setHeight(short height) {
            this.height = height;
        }

        public float getHeightInPoints() {
            return heightInPoints;
        }

        public void setHeightInPoints(float heightInPoints) {
            this.heightInPoints = heightInPoints;
        }

        public CellStyle getRowStyle() {
            return rowStyle;
        }

        public void setRowStyle(CellStyle rowStyle) {
            this.rowStyle = rowStyle;
        }

        public boolean isZeroHeight() {
            return zeroHeight;
        }

        public void setZeroHeight(boolean zeroHeight) {
            this.zeroHeight = zeroHeight;
        }

        public InnerCell getCell(Integer colIndex) {
            return cellMap.get(colIndex);
        }

        public void setCell(Integer colIndex, InnerCell cell) {
            cellMap.put(colIndex, cell);
        }
    }

    private static class InnerCell {

        private CellStyle cellStyle;
        private int       cellType;

        public CellStyle getCellStyle() {
            return cellStyle;
        }

        public void setCellStyle(CellStyle cellStyle) {
            this.cellStyle = cellStyle;
        }

        public int getCellType() {
            return cellType;
        }

        public void setCellType(int cellType) {
            this.cellType = cellType;
        }
    }

    @SuppressWarnings("unchecked")
    private static void write(boolean useTemplate, Workbook workbook, OutputStream outputStream,
                              ExcelWriteSheetProcessor<?>... sheetProcessors) {

        for (@SuppressWarnings("rawtypes")
        ExcelWriteSheetProcessor sheetProcessor : sheetProcessors) {
            @SuppressWarnings("rawtypes")
            ExcelWriteContext context = new ExcelWriteContext();

            try {
                if (sheetProcessor == null) {
                    continue;
                }
                String sheetName = sheetProcessor.getSheetName();
                Integer sheetIndex = sheetProcessor.getSheetIndex();
                Sheet sheet = null;
                // sheetName priority,
                if (useTemplate) {
                    if (sheetName != null) {
                        sheet = workbook.getSheet(sheetName);
                        if (sheet != null && sheetIndex != null && !sheetIndex.equals(workbook.getSheetIndex(sheet))) {
                            throw new IllegalArgumentException("sheetName[" + sheetName + "] and sheetIndex["
                                                               + sheetIndex + "] not match.");
                        }
                    } else if (sheetIndex != null) {
                        sheet = workbook.getSheetAt(sheetIndex);
                    } else {
                        throw new IllegalArgumentException("sheetName or sheetIndex can't be null");
                    }
                    if (sheet == null) {
                        throw new IllegalArgumentException("Sheet Not Found Exception. for sheet name:" + sheetName);
                    }
                } else {
                    if (sheetName != null) {
                        sheet = workbook.getSheet(sheetName);
                        if (sheet != null) {
                            if (sheetIndex != null && !sheetIndex.equals(workbook.getSheetIndex(sheet))) {
                                throw new IllegalArgumentException("sheetName[" + sheetName + "] and sheetIndex["
                                                                   + sheetIndex + "] not match.");
                            }
                        } else {
                            sheet = workbook.createSheet(sheetName);
                            if (sheetIndex != null) {
                                workbook.setSheetOrder(sheetName, sheetIndex);
                            }
                        }
                    } else if (sheetIndex != null) {
                        sheet = workbook.createSheet();
                        workbook.setSheetOrder(sheet.getSheetName(), sheetIndex);
                    } else {
                        throw new IllegalArgumentException("sheetName or sheetIndex can't be null");
                    }
                }

                if (sheetIndex == null) {
                    sheetIndex = workbook.getSheetIndex(sheet);
                }
                if (sheetName == null) {
                    sheetName = sheet.getSheetName();
                }
                // proc sheet
                context.setCurSheet(sheet);
                context.setCurSheetIndex(sheetIndex);
                context.setCurSheetName(sheet.getSheetName());
                context.setCurRow(null);
                context.setCurRowIndex(null);
                context.setCurCell(null);
                context.setCurColIndex(null);
                // beforeProcess
                sheetProcessor.beforeProcess(context);

                InnerRow templateRow = null;
                if (sheetProcessor.getTemplateRowIndex() != null) {
                    Row tempRow = sheet.getRow(sheetProcessor.getTemplateRowIndex());
                    if (tempRow != null) {
                        templateRow = new InnerRow();
                        templateRow.setHeight(tempRow.getHeight());
                        templateRow.setHeightInPoints(tempRow.getHeightInPoints());
                        templateRow.setRowStyle(tempRow.getRowStyle());
                        templateRow.setZeroHeight(tempRow.getZeroHeight());
                        for (int i = tempRow.getFirstCellNum(); i <= tempRow.getLastCellNum(); i++) {
                            Cell cell = tempRow.getCell(i);
                            if (cell != null) {
                                InnerCell innerCell = new InnerCell();
                                innerCell.setCellStyle(cell.getCellStyle());
                                innerCell.setCellType(cell.getCellType());
                                templateRow.setCell(i, innerCell);
                            }
                        }
                    }
                }

                int writeRowIndex = sheetProcessor.getRowStartIndex();
                for (@SuppressWarnings("rawtypes")
                List dataList = sheetProcessor.getDataList(context); //
                dataList != null && !dataList.isEmpty(); //
                dataList = sheetProcessor.getDataList(context)) {
                    for (Object rowData : dataList) {
                        // proc row
                        Row row = sheet.getRow(writeRowIndex);
                        if (row == null) {
                            row = sheet.createRow(writeRowIndex);
                        }
                        if (templateRow != null) {
                            row.setHeight(templateRow.getHeight());
                            row.setHeightInPoints(templateRow.getHeightInPoints());
                            row.setRowStyle(templateRow.getRowStyle());
                            row.setZeroHeight(templateRow.isZeroHeight());
                        }
                        context.setCurRow(row);
                        context.setCurRowIndex(writeRowIndex);
                        context.setCurColIndex(null);
                        context.setCurCell(null);
                        // ///////
                        if (rowData != null) {
                            writeRow(context, templateRow, row, rowData, sheetProcessor);
                        }

                        Row reRow = null;
                        try {
                            reRow = sheetProcessor.process(context, rowData, row);
                        } catch (RuntimeException e) {
                            if (e instanceof ExcelWriteException) {
                                ExcelWriteException ewe = (ExcelWriteException) e;
                                // ef.setColIndex(null); //user may want to set
                                // this value,
                                ewe.setRowIndex(writeRowIndex);
                                throw ewe;
                            } else {
                                ExcelWriteException ewe = new ExcelWriteException(e);
                                ewe.setColIndex(null);
                                ewe.setCode(ExcelWriteException.CODE_OF_PROCESS_EXCEPTION);
                                ewe.setRowIndex(writeRowIndex);
                                throw ewe;
                            }
                        }

                        if (reRow == null) {
                            sheet.removeRow(row);
                        } else {
                            writeRowIndex++;
                        }
                    }
                }
                if (templateRow != null) {
                    writeDataValidations(sheet, sheetProcessor);
                }
            } catch (RuntimeException e) {
                sheetProcessor.onException(context, e);
            } finally {
                sheetProcessor.afterProcess(context);
            }
        }

        try {
            workbook.write(outputStream);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    @SuppressWarnings("rawtypes")
    private static void writeDataValidations(Sheet sheet, ExcelWriteSheetProcessor sheetProcessor) {
        int templateRowIndex = sheetProcessor.getTemplateRowIndex();
        int rowStartIndex = sheetProcessor.getRowStartIndex();

        Set<Integer> configColIndexSet = new HashSet<Integer>();
        for (Entry<String, Map<Integer, InnerWriteCellProcessorWrapper>> fieldIndexMapping : sheetProcessor.getFieldMapping().entrySet()) {
            if (fieldIndexMapping == null || fieldIndexMapping.getValue() == null) {
                continue;
            }
            for (Entry<Integer, InnerWriteCellProcessorWrapper> indexProcessorMapping : fieldIndexMapping.getValue().entrySet()) {
                if (indexProcessorMapping == null || indexProcessorMapping.getKey() == null) {
                    continue;
                }
                configColIndexSet.add(indexProcessorMapping.getKey());
            }
        }

        List<? extends DataValidation> dataValidations = sheet.getDataValidations();
        if (dataValidations != null) {
            for (DataValidation dataValidation : dataValidations) {
                if (dataValidation == null) {
                    continue;
                }
                CellRangeAddressList cellRangeAddressList = dataValidation.getRegions();
                if (cellRangeAddressList == null) {
                    continue;
                }

                CellRangeAddress[] cellRangeAddresses = cellRangeAddressList.getCellRangeAddresses();
                if (cellRangeAddresses == null || cellRangeAddresses.length == 0) {
                    continue;
                }

                CellRangeAddressList newCellRangeAddressList = new CellRangeAddressList();
                boolean validationContains = false;
                for (CellRangeAddress cellRangeAddress : cellRangeAddresses) {
                    if (cellRangeAddress == null) {
                        continue;
                    }
                    if (templateRowIndex < cellRangeAddress.getFirstRow()
                        || templateRowIndex > cellRangeAddress.getLastRow()) {// specify
                                                                              // row
                        continue;
                    }
                    for (Integer configColIndex : configColIndexSet) {
                        if (configColIndex < cellRangeAddress.getFirstColumn()
                            || configColIndex > cellRangeAddress.getLastColumn()) {// specify column
                            continue;
                        }
                        newCellRangeAddressList.addCellRangeAddress(rowStartIndex, configColIndex,
                                                                    sheet.getLastRowNum(), configColIndex);
                        validationContains = true;
                    }
                }
                if (validationContains) {
                    DataValidation newDataValidation = sheet.getDataValidationHelper().createValidation(dataValidation.getValidationConstraint(),
                                                                                                        newCellRangeAddressList);
                    sheet.addValidationData(newDataValidation);
                }
            }
        }
    }

    @SuppressWarnings({ "rawtypes", "unchecked" })
    private static void writeRow(ExcelWriteContext context, InnerRow templateRow, Row row, Object rowData,
                                 ExcelWriteSheetProcessor sheetProcessor) {

        boolean useTemplate = false;
        if (templateRow != null) {
            useTemplate = true;
        }
        Class<?> clazz = rowData.getClass();
        ExcelWriteFieldMapping fieldMapping = sheetProcessor.getFieldMapping();
        for (Entry<String, Map<Integer, InnerWriteCellProcessorWrapper>> entry : fieldMapping.entrySet()) {
            String fieldName = entry.getKey();
            Map<Integer, InnerWriteCellProcessorWrapper> map = entry.getValue();
            for (Map.Entry<Integer, InnerWriteCellProcessorWrapper> fieldValueMapping : map.entrySet()) {
                Integer colIndex = fieldValueMapping.getKey();
                InnerWriteCellProcessorWrapper cellProcessorWrapper = fieldValueMapping.getValue();

                PropertyDescriptor pd = BeanUtils.getPropertyDescriptor(clazz, fieldName);
                if (pd.getReadMethod() == null) {
                    continue;
                }
                Object val = null;
                try {
                    val = pd.getReadMethod().invoke(rowData, (Object[]) null);
                } catch (Exception e) {
                    throw new RuntimeException(e);
                }

                // proc cell
                Cell cell = row.getCell(colIndex);
                if (cell == null) {
                    cell = row.createCell(colIndex);
                }
                if (templateRow != null) {
                    InnerCell tempalteCell = templateRow.getCell(colIndex);
                    if (tempalteCell != null) {
                        cell.setCellStyle(tempalteCell.getCellStyle());
                        cell.setCellType(tempalteCell.getCellType());
                    }
                }
                context.setCurColIndex(colIndex);
                context.setCurCell(cell);

                ExcelWriteCellValueMapping valueMapping = cellProcessorWrapper.getValueMapping();
                ExcelWriteCellProcessor processor = cellProcessorWrapper.getProcessor();
                if (valueMapping != null) {
                    String key = null;
                    if (val != null) {
                        key = val.toString();
                    }
                    Object cval = valueMapping.get(key);
                    if (cval != null) {
                        writeCell(row.getRowNum(), colIndex, cell, cval, useTemplate);
                    } else {
                        if (!valueMapping.containsKey(key)) {
                            if (valueMapping.isSetDefaultValue()) {
                                if (valueMapping.isSetDefaultValueWithDefaultInput()) {
                                    writeCell(row.getRowNum(), colIndex, cell, val, useTemplate);
                                } else {
                                    writeCell(row.getRowNum(), colIndex, cell, valueMapping.getDefaultValue(),
                                              useTemplate);
                                }
                            } else if (valueMapping.getDefaultProcessor() != null) {
                                cell = valueMapping.getDefaultProcessor().process(context, rowData, cell);
                            } else {
                                ExcelWriteException ex = new ExcelWriteException();
                                ex.setCode(ExcelWriteException.CODE_OF_FIELD_VALUE_NOT_MATCHED);
                                ex.setColIndex(colIndex);
                                ex.setRowIndex(row.getRowNum());
                                throw ex;
                            }
                        } else {
                            // contains null
                            // ok
                        }
                    }
                } else if (processor != null) {
                    writeCell(cell, val, useTemplate);
                    try {
                        cell = processor.process(context, rowData, cell);
                    } catch (RuntimeException e) {
                        if (e instanceof ExcelWriteException) {
                            ExcelWriteException ewe = (ExcelWriteException) e;
                            ewe.setColIndex(colIndex);
                            ewe.setRowIndex(row.getRowNum());
                            throw ewe;
                        } else {
                            ExcelWriteException ewe = new ExcelWriteException(e);
                            ewe.setColIndex(colIndex);
                            ewe.setCode(ExcelWriteException.CODE_OF_PROCESS_EXCEPTION);
                            ewe.setRowIndex(row.getRowNum());
                            throw ewe;
                        }
                    }
                } else {
                    writeCell(cell, val, useTemplate);
                }

                if (cell == null) {
                    row.removeCell(cell);
                }
            }
        }
    }

    private static void writeCell(int rowIndex, int colIndex, Cell cell, Object val, boolean userTemplate) {
        try {
            writeCell(cell, val, userTemplate);
        } catch (RuntimeException e) {
            ExcelWriteException ewe = new ExcelWriteException(e);
            ewe.setColIndex(colIndex);
            ewe.setCode(ExcelWriteException.CODE_OF_PROCESS_EXCEPTION);
            ewe.setRowIndex(rowIndex);
            throw ewe;
        }
    }

    public static void writeCell(Cell cell, Object val) {
        if (cell.getCellStyle() != null && cell.getCellStyle().getDataFormat() > 0) {
            writeCell(cell, val, true);
        } else {
            writeCell(cell, val, false);
        }
    }

    @SuppressWarnings("unused")
    private static void writeCell(Cell cell, Object val, boolean userTemplate) {
        if (val == null) {
            cell.setCellValue((String) null);
            return;
        }
        Class<?> clazz = val.getClass();
        if (val instanceof Byte) {// Double
            Byte temp = (Byte) val;
            cell.setCellValue((double) temp.byteValue());
        } else if (val instanceof Short) {
            Short temp = (Short) val;
            cell.setCellValue((double) temp.shortValue());
        } else if (val instanceof Integer) {
            Integer temp = (Integer) val;
            cell.setCellValue((double) temp.intValue());
        } else if (val instanceof Long) {
            Long temp = (Long) val;
            cell.setCellValue((double) temp.longValue());
        } else if (val instanceof Float) {
            Float temp = (Float) val;
            cell.setCellValue((double) temp.floatValue());
        } else if (val instanceof Double) {
            Double temp = (Double) val;
            cell.setCellValue((double) temp.doubleValue());
        } else if (val instanceof Date) {// Date
            Date dateVal = (Date) val;
            long time = dateVal.getTime();
            // read is based on 1899/12/31 but DateUtil.getExcelDate is base on
            // 1900/01/01
            if (time >= TIME_1899_12_31_00_00_00_000 && time < TIME_1900_01_01_00_00_00_000) {
                Date incOneDay = new Date(time + 24 * 60 * 60 * 1000);
                double d = DateUtil.getExcelDate(incOneDay);
                cell.setCellValue(d - 1);
            } else {
                cell.setCellValue(dateVal);
            }

            if (!userTemplate) {
                Workbook wb = cell.getRow().getSheet().getWorkbook();
                CellStyle cellStyle = cell.getCellStyle();
                if (cellStyle == null) {
                    cellStyle = wb.createCellStyle();
                }
                DataFormat dataFormat = wb.getCreationHelper().createDataFormat();
                // @see #BuiltinFormats
                // 0xe, "m/d/yy"
                // 0x14 "h:mm"
                // 0x16 "m/d/yy h:mm"
                // {@linke https://en.wikipedia.org/wiki/Year_10,000_problem}
                /** [1899/12/31 00:00:00:000~1900/01/01 00:00:000) */
                if (time >= TIME_1899_12_31_00_00_00_000 && time < TIME_1900_01_02_00_00_00_000) {
                    cellStyle.setDataFormat(dataFormat.getFormat("h:mm"));
                    // cellStyle.setDataFormat(dataFormat.getFormat("m/d/yy h:mm"));
                } else {
                    // if ( time % (24 * 60 * 60 * 1000) == 0) {//for time
                    // zone,we can't use this way.
                    Calendar calendar = Calendar.getInstance();
                    calendar.setTime(dateVal);
                    int hour = calendar.get(Calendar.HOUR_OF_DAY);
                    int minute = calendar.get(Calendar.MINUTE);
                    int second = calendar.get(Calendar.SECOND);
                    int millisecond = calendar.get(Calendar.MILLISECOND);
                    if (millisecond == 0 && second == 0 && minute == 0 && hour == 0) {
                        cellStyle.setDataFormat(dataFormat.getFormat("m/d/yy"));
                    } else {
                        cellStyle.setDataFormat(dataFormat.getFormat("m/d/yy h:mm"));
                    }
                }
                cell.setCellStyle(cellStyle);
            }
        } else if (val instanceof Boolean) {// Boolean
            cell.setCellValue(((Boolean) val).booleanValue());
        } else {// String
            cell.setCellValue((String) val.toString());
        }
    }

    /**
     * Convert excel column character index (such as 'A','B','AC') to integer index (0-based)
     * note: character index ignores case
     * eg: 'A'  -> 0
     *     'B'  -> 1
     *     'AC' -> 28
     *     'aC' -> 28
     *     'Ac' -> 28
     * @param colIndex column character index
     * @return column integer index
     * @see #convertColIntIndexToCharIndex
     */
    public static int convertColCharIndexToIntIndex(String colIndex) {
        char[] chars = colIndex.toCharArray();
        int index = 0;
        int baseStep = 'z' - 'a' + 1;
        int curStep = 1;
        for (int i = chars.length - 1; i >= 0; i--) {
            char ch = chars[i];
            if (ch >= 'A' && ch <= 'Z') {
                index += (ch - 'A' + 1) * curStep;
            } else if (ch >= 'a' && ch <= 'z') {
                index += (ch - 'a' + 1) * curStep;
            } else {
                throw new IllegalArgumentException("colIndex must be a-z or A-Z,unexpected character:" + ch);
            }
            curStep *= baseStep;
        }
        index--;
        return index;
    }

    /**
     * Convert excel column integer index (0-based) to character index (such as 'A','B','AC')
     * eg: 0  -> 'A'
     *     1  -> 'B'
     *     28 -> 'AC'
     * @param colIndex column integer index.
     * @return column character index in capitals
     * @ #convertColCharIndexToIntIndex
     */
    public static String convertColIntIndexToCharIndex(Integer index) {
        Assert.isTrue(index >= 0);
        StringBuilder sb = new StringBuilder();
        do {
            char c = (char) ((index % 26) + 'A');
            sb.insert(0, c);
            index = index / 26 - 1;
        } while (index >= 0);
        return sb.toString();
    }
}
