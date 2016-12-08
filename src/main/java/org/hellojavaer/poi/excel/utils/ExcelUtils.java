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

import com.alibaba.fastjson.util.TypeUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hellojavaer.poi.excel.utils.common.Assert;
import org.hellojavaer.poi.excel.utils.read.*;
import org.hellojavaer.poi.excel.utils.read.ExcelReadFieldMapping.ExcelReadFieldMappingAttribute;
import org.hellojavaer.poi.excel.utils.write.*;
import org.hellojavaer.poi.excel.utils.write.ExcelWriteFieldMapping.ExcelWriteFieldMappingAttribute;
import org.springframework.beans.BeanUtils;

import java.beans.PropertyDescriptor;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Modifier;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.Map.Entry;

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

    private static void convertFieldMapping(Sheet sheet, ExcelReadSheetProcessor<?> sheetProcessor,
                                            Map<String, Map<String, ExcelReadFieldMappingAttribute>> src,
                                            Map<Integer, Map<String, ExcelReadFieldMappingAttribute>> tar) {
        if (src == null) {
            return;
        }
        Integer headRowIndex = sheetProcessor.getHeadRowIndex();
        Map<String, Integer> colCache = new HashMap<String, Integer>();
        if (headRowIndex != null) {
            Row row = sheet.getRow(headRowIndex);
            if (row != null) {
                int start = row.getFirstCellNum();
                int end = row.getLastCellNum();
                for (int i = start; i < end; i++) {
                    Cell cell = row.getCell(i);
                    Object cellValue = _readCell(cell);
                    if (cellValue != null) {
                        String strVal = cellValue.toString().trim();
                        colCache.put(strVal, i);
                    }
                }
            }
        }

        for (Map.Entry<String, Map<String, ExcelReadFieldMappingAttribute>> entry : src.entrySet()) {
            String colIndexOrColName = entry.getKey();
            Integer colIndex = null;
            if (headRowIndex == null) {
                colIndex = convertColCharIndexToIntIndex(colIndexOrColName);
            } else {
                colIndex = colCache.get(colIndexOrColName);
                if (colIndex == null) {
                    throw new IllegalStateException("For sheet:" + sheet.getSheetName() + " headRowIndex:"
                                                    + headRowIndex + " can't find colum named '" + colIndexOrColName
                                                    + "'");
                }
            }
            tar.put(colIndex, entry.getValue());
        }
    }

    private static void readConfigParamVerify(ExcelReadSheetProcessor<?> sheetProcessor,
                                              Map<Integer, Map<String, ExcelReadFieldMappingAttribute>> fieldMapping) {
        Class<?> clazz = sheetProcessor.getTargetClass();

        for (Entry<Integer, Map<String, ExcelReadFieldMappingAttribute>> indexFieldMapping : fieldMapping.entrySet()) {
            for (Map.Entry<String, ExcelReadFieldMappingAttribute> filedMapping : indexFieldMapping.getValue().entrySet()) {
                String fieldName = filedMapping.getKey();
                if (fieldName != null) {
                    PropertyDescriptor pd = getPropertyDescriptor(clazz, fieldName);
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
                ExcelReadContext context = new ExcelReadContext();
                String sheetName = sheetProcessor.getSheetName();
                Integer sheetIndex = sheetProcessor.getSheetIndex();
                try {
                    Class clazz = sheetProcessor.getTargetClass();
                    context.setCurSheetIndex(sheetIndex);
                    context.setCurSheetName(sheetName);
                    Sheet sheet = null;
                    if (sheetName != null) {
                        try {
                            sheet = workbook.getSheet(sheetName);
                        } catch (IllegalArgumentException e) {
                            // ignore
                        }
                        if (sheet != null && sheetIndex != null && !sheetIndex.equals(workbook.getSheetIndex(sheet))) {
                            ExcelReadException e = new ExcelReadException();
                            e.setCode(ExcelReadException.CODE_OF_SHEET_NAME_AND_INDEX_NOT_MATCH);
                            throw e;
                        }
                    } else if (sheetIndex != null) {
                        try {
                            sheet = workbook.getSheetAt(sheetIndex);
                        } catch (IllegalArgumentException e) {
                            // ignore
                        }
                    } else {
                        ExcelReadException e = new ExcelReadException();
                        e.setCode(ExcelReadException.CODE_OF_SHEET_NAME_AND_INDEX_IS_EMPTY);
                        throw e;
                    }
                    if (sheet == null) {
                        ExcelReadException e = new ExcelReadException("name of '" + sheetName + "' sheet is not exist"
                                                                      + sheetName);
                        e.setCode(ExcelReadException.CODE_OF_SHEET_NOT_EXIST);
                        throw e;
                    }

                    if (sheetIndex == null) {
                        sheetIndex = workbook.getSheetIndex(sheet);
                    }
                    if (sheetName == null) {
                        sheetName = sheet.getSheetName();
                    }
                    // do check
                    Map<Integer, Map<String, ExcelReadFieldMappingAttribute>> fieldMapping = new HashMap<Integer, Map<String, ExcelReadFieldMappingAttribute>>();
                    Map<String, Map<String, ExcelReadFieldMappingAttribute>> src = null;
                    if (sheetProcessor.getFieldMapping() != null) {
                        src = sheetProcessor.getFieldMapping().export();
                    }
                    convertFieldMapping(sheet, sheetProcessor, src, fieldMapping);
                    if (sheetProcessor.getTargetClass() != null && sheetProcessor.getFieldMapping() != null
                        && !Map.class.isAssignableFrom(sheetProcessor.getTargetClass())) {
                        readConfigParamVerify(sheetProcessor, fieldMapping);
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
                    int startRow = sheetProcessor.getStartRowIndex();
                    Integer rowEndIndex = sheetProcessor.getEndRowIndex();
                    int actLastRow = sheet.getLastRowNum();
                    if (rowEndIndex != null) {
                        if (rowEndIndex > actLastRow) {
                            rowEndIndex = actLastRow;
                        }
                    } else {
                        rowEndIndex = actLastRow;
                    }

                    ExcelProcessControllerImpl controller = new ExcelProcessControllerImpl();
                    if (pageSize != null) {
                        int total = rowEndIndex - startRow + 1;
                        int pageCount = (total + pageSize - 1) / pageSize;
                        for (int i = 0; i < pageCount; i++) {
                            int start = startRow + pageSize * i;
                            int size = pageSize;
                            if (i == pageCount - 1) {
                                size = rowEndIndex - start + 1;
                            }
                            read(controller, context, sheet, start, size, fieldMapping, clazz,
                                 sheetProcessor.getRowProcessor(), sheetProcessor.isTrimSpace());
                            sheetProcessor.process(context, context.getDataList());
                            context.getDataList().clear();
                            if (controller.isDoBreak()) {
                                controller.reset();
                                break;
                            }
                        }
                    } else {
                        read(controller, context, sheet, startRow, rowEndIndex - startRow + 1, fieldMapping, clazz,
                             sheetProcessor.getRowProcessor(), sheetProcessor.isTrimSpace());
                        sheetProcessor.process(context, context.getDataList());
                        context.getDataList().clear();
                    }
                } catch (Throwable e) {
                    if (e instanceof ExcelReadException) {
                        ExcelReadException e0 = (ExcelReadException) e;
                        e0.setSheetName(sheetName);
                        e0.setSheetIndex(sheetIndex);
                        sheetProcessor.onException(context, (ExcelReadException) e0);
                    } else {
                        ExcelReadException e0 = new ExcelReadException(e);
                        e0.setSheetName(sheetName);
                        e0.setSheetIndex(sheetIndex);
                        e0.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                        sheetProcessor.onException(context, e0);
                    }
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

    private static <T> void read(ExcelProcessControllerImpl controller, ExcelReadContext<T> context, Sheet sheet,
                                 int startRow, Integer pageSize,
                                 Map<Integer, Map<String, ExcelReadFieldMappingAttribute>> fieldMapping,
                                 Class<T> targetClass, ExcelReadRowProcessor<T> processor, boolean isTrimSpace) {
        Assert.isTrue(sheet != null, "sheet can't be null");
        Assert.isTrue(startRow >= 0, "startRow must greater than or equal to 0");
        Assert.isTrue(pageSize == null || pageSize >= 1, "pageSize == null || pageSize >= 1");
        Assert.isTrue(fieldMapping != null, "fieldMapping can't be null");
        // Assert.isTrue(targetClass != null, "clazz can't be null");

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

            T t = null;
            if (!fieldMapping.isEmpty()) {
                t = readRow(context, row, fieldMapping, targetClass, processor, isTrimSpace);
            }
            if (processor != null) {
                try {
                    controller.reset();
                    t = processor.process(controller, context, row, t);
                } catch (Throwable e) {
                    if (e instanceof ExcelReadException) {
                        ExcelReadException e0 = (ExcelReadException) e;
                        e0.setRowIndex(row.getRowNum());
                        // e0.setColIndex(null);//user may want to set this value,
                        throw e0;
                    } else {
                        ExcelReadException e0 = new ExcelReadException(e);
                        e0.setRowIndex(row.getRowNum());
                        e0.setColIndex(null);
                        e0.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                        throw e0;
                    }
                }
            }
            if (!controller.isDoSkip()) {
                list.add(t);
            }
            if (controller.isDoBreak()) {
                break;
            }
        }
    }

    @SuppressWarnings({ "unchecked", "rawtypes" })
    private static <T> T readRow(ExcelReadContext<T> context, Row row,
                                 Map<Integer, Map<String, ExcelReadFieldMappingAttribute>> fieldMapping,
                                 Class<T> targetClass, ExcelReadRowProcessor<T> processor, boolean isTrimSpace) {
        try {
            context.setCurRowData(targetClass.newInstance());
        } catch (Exception e1) {
            throw new RuntimeException(e1);
        }
        int curRowIndex = context.getCurRowIndex();
        for (Entry<Integer, Map<String, ExcelReadFieldMappingAttribute>> fieldMappingEntry : fieldMapping.entrySet()) {
            int curColIndex = fieldMappingEntry.getKey();// excel index;
            // proc cell
            context.setCurColIndex(curColIndex);

            Cell cell = null;
            if (row != null) {
                cell = row.getCell(curColIndex);
            }
            context.setCurCell(cell);

            Map<String, ExcelReadFieldMappingAttribute> fields = fieldMappingEntry.getValue();
            for (Map.Entry<String, ExcelReadFieldMappingAttribute> fieldEntry : fields.entrySet()) {
                String fieldName = fieldEntry.getKey();
                ExcelReadFieldMappingAttribute attribute = fieldEntry.getValue();
                // proccess link
                String linkField = attribute.getLinkField();
                if (linkField != null) {
                    String address = null;
                    if (cell != null) {
                        Hyperlink hyperlink = cell.getHyperlink();
                        if (hyperlink != null) {
                            address = hyperlink.getAddress();
                        }
                    }
                    if (isTrimSpace && address != null) {
                        address = address.trim();
                        if (address.length() == 0) {
                            address = null;
                        }
                    }
                    if (Map.class.isAssignableFrom(targetClass)) {// map
                        ((Map) context.getCurRowData()).put(linkField, address);
                    } else {// java bean
                        try {
                            setProperty(context.getCurRowData(), linkField, address);
                        } catch (Throwable e) {
                            if (e instanceof ExcelReadException) {
                                ExcelReadException e0 = (ExcelReadException) e;
                                e0.setRowIndex(curRowIndex);
                                e0.setColIndex(curColIndex);
                                throw e0;
                            } else {
                                ExcelReadException e0 = new ExcelReadException(e);
                                e0.setRowIndex(curRowIndex);
                                e0.setColIndex(curColIndex);
                                e0.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                                throw e0;
                            }
                        }
                    }
                }

                Object value = _readCell(cell);
                if (value != null && value instanceof String && isTrimSpace) {
                    value = ((String) value).trim();
                    if (((String) value).length() == 0) {
                        value = null;
                    }
                }
                if (value == null && attribute.isRequired()) {
                    ExcelReadException e = new ExcelReadException("Cell value is null");
                    e.setRowIndex(curRowIndex);
                    e.setColIndex(curColIndex);
                    e.setCode(ExcelReadException.CODE_OF_CELL_VALUE_REQUIRED);
                    throw e;
                }
                //
                try {
                    if (Map.class.isAssignableFrom(targetClass)) {// map
                        value = procValueConvert(context, row, cell, attribute, fieldName, value);
                        ((Map) context.getCurRowData()).put(fieldName, value);
                    } else {// java bean
                        value = procValueConvert(context, row, cell, attribute, fieldName, value);
                        setProperty(context.getCurRowData(), fieldName, value);
                    }
                } catch (Throwable e) {
                    if (e instanceof ExcelReadException) {
                        ExcelReadException e0 = (ExcelReadException) e;
                        e0.setRowIndex(curRowIndex);
                        e0.setColIndex(curColIndex);
                        throw e0;
                    } else {
                        ExcelReadException e0 = new ExcelReadException(e);
                        e0.setRowIndex(curRowIndex);
                        e0.setColIndex(curColIndex);
                        e0.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                        throw e0;
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
        Object val = _readCell(cell);
        return new ExcelCellValue(val);
    }

    private static Object _readCell(Cell cell) {
        if (cell == null) {
            return null;
        }
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
                ExcelReadException e = new ExcelReadException("Cell type error");
                e.setRowIndex(cell.getRowIndex());
                e.setColIndex(cell.getColumnIndex());
                e.setCode(ExcelReadException.CODE_OF_CELL_ERROR);
                throw e;
            case Cell.CELL_TYPE_FORMULA:
                value = cell.getCellFormula();
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
                value = cell.getStringCellValue();
                break;
            default:
                throw new RuntimeException("unsupport cell type " + cellType);
        }
        return value;
    }

    private static Object procValueConvert(ExcelReadContext<?> context, Row row, Cell cell,
                                           ExcelReadFieldMappingAttribute entry, String fieldName, Object value) {
        Object convertedValue = value;
        if (entry.getValueMapping() != null) {
            ExcelReadCellValueMapping valueMapping = entry.getValueMapping();
            String strValue = TypeUtils.castToString(value);
            convertedValue = valueMapping.get(strValue);
            if (convertedValue == null) {
                if (!valueMapping.containsKey(strValue)) {
                    if (valueMapping.isSettedDefaultValue()) {
                        if (valueMapping.isSettedDefaultValueWithDefaultInput()) {
                            convertedValue = value;
                        } else {
                            convertedValue = valueMapping.getDefaultValue();
                        }
                    } else if (valueMapping.getDefaultProcessor() != null) {
                        try {
                            convertedValue = valueMapping.getDefaultProcessor().process(context, cell,
                                                                                        new ExcelCellValue(value));
                        } catch (Throwable e) {
                            if (e instanceof ExcelReadException) {
                                ExcelReadException e0 = (ExcelReadException) e;
                                e0.setRowIndex(row.getRowNum());
                                e0.setColIndex(cell.getColumnIndex());
                                throw e0;
                            } else {
                                ExcelReadException e0 = new ExcelReadException(e);
                                e0.setRowIndex(row.getRowNum());
                                e0.setColIndex(cell.getColumnIndex());
                                e0.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                                throw e0;
                            }
                        }
                        if (convertedValue != null && convertedValue instanceof ExcelCellValue) {
                            convertedValue = value;
                        }
                    } else {
                        ExcelReadException e = new ExcelReadException("Cell value is value " + strValue);
                        e.setRowIndex(row.getRowNum());
                        e.setColIndex(cell.getColumnIndex());
                        e.setCode(ExcelReadException.CODE_OF_CELL_VALUE_NOT_MATCH);
                        throw e;
                    }
                }
            }
        } else if (entry.getCellProcessor() != null) {
            try {
                convertedValue = entry.getCellProcessor().process(context, cell, new ExcelCellValue(value));
            } catch (Throwable e) {
                if (e instanceof ExcelReadException) {
                    ExcelReadException e0 = (ExcelReadException) e;
                    e0.setRowIndex(row.getRowNum());
                    e0.setColIndex(cell.getColumnIndex());
                    throw e0;
                } else {
                    ExcelReadException e0 = new ExcelReadException(e);
                    e0.setRowIndex(row.getRowNum());
                    e0.setColIndex(cell.getColumnIndex());
                    e0.setCode(ExcelReadException.CODE_OF_PROCESS_EXCEPTION);
                    throw e0;
                }
            }
            if (convertedValue != null && convertedValue instanceof ExcelCellValue) {
                convertedValue = value;
            }
        }
        if (convertedValue == null && entry.isRequired()) {
            ExcelReadException e = new ExcelReadException("Cell value is null");
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

    @SuppressWarnings("rawtypes")
    private static void writeHead(boolean useTemplate, Sheet sheet, ExcelWriteSheetProcessor sheetProcessor) {
        Integer headRowIndex = sheetProcessor.getHeadRowIndex();
        if (headRowIndex == null) {
            return;
        }
        Workbook wookbook = sheet.getWorkbook();
        // use theme
        CellStyle style = null;
        if (!useTemplate && sheetProcessor.getTheme() != null) {
            int theme = sheetProcessor.getTheme();
            if (theme == ExcelWriteTheme.BASE) {
                style = wookbook.createCellStyle();
                style.setFillPattern(CellStyle.SOLID_FOREGROUND);
                style.setFillForegroundColor((short) 44);
                style.setBorderBottom(CellStyle.BORDER_THIN);
                style.setBorderLeft(CellStyle.BORDER_THIN);
                style.setBorderRight(CellStyle.BORDER_THIN);
                style.setBorderTop(CellStyle.BORDER_THIN);
                // style.setBottomBorderColor((short) 44);
                style.setAlignment(CellStyle.ALIGN_CENTER);
            }
            // freeze Pane
            if (sheetProcessor.getHeadRowIndex() != null && sheetProcessor.getHeadRowIndex() == 0) {
                sheet.createFreezePane(0, 1, 0, 1);
            }
        }

        Row row = sheet.getRow(headRowIndex);
        if (row == null) {
            row = sheet.createRow(headRowIndex);
        }
        for (Map.Entry<String, Map<Integer, ExcelWriteFieldMappingAttribute>> entry : sheetProcessor.getFieldMapping().export().entrySet()) {
            Map<Integer, ExcelWriteFieldMappingAttribute> map = entry.getValue();
            if (map != null) {
                for (Map.Entry<Integer, ExcelWriteFieldMappingAttribute> entry2 : map.entrySet()) {
                    String head = entry2.getValue().getHead();
                    Integer colIndex = entry2.getKey();
                    Cell cell = row.getCell(colIndex);
                    if (cell == null) {
                        cell = row.createCell(colIndex);
                    }
                    // use theme
                    if (!useTemplate && sheetProcessor.getTheme() != null) {
                        cell.setCellStyle(style);

                    }
                    cell.setCellValue(head);
                }
            }
        }

    }

    private static void write(boolean useTemplate, Workbook workbook, OutputStream outputStream,
                              ExcelWriteSheetProcessor<?>... sheetProcessors) {

        for (ExcelWriteSheetProcessor sheetProcessor : sheetProcessors) {
            ExcelWriteContext context = new ExcelWriteContext();
            String sheetName = sheetProcessor.getSheetName();
            Integer sheetIndex = sheetProcessor.getSheetIndex();
            try {
                if (sheetProcessor == null) {
                    continue;
                }
                Sheet sheet = null;
                if (sheetProcessor.getTemplateStartRowIndex() == null
                    && sheetProcessor.getTemplateEndRowIndex() == null) {
                    sheetProcessor.setTemplateRows(sheetProcessor.getStartRowIndex(), sheetProcessor.getStartRowIndex());
                }
                // sheetName priority,
                if (useTemplate) {
                    if (sheetName != null) {
                        try {
                            sheet = workbook.getSheet(sheetName);
                        } catch (IllegalArgumentException e) {
                            // ignore
                        }
                        if (sheet != null && sheetIndex != null && !sheetIndex.equals(workbook.getSheetIndex(sheet))) {
                            ExcelWriteException e = new ExcelWriteException("sheetName[" + sheetName
                                                                            + "] and sheetIndex[" + sheetIndex
                                                                            + "] not match.");
                            e.setCode(ExcelWriteException.CODE_OF_SHEET_NAME_AND_INDEX_NOT_MATCH);
                            throw e;
                        }
                    } else if (sheetIndex != null) {
                        try {
                            sheet = workbook.getSheetAt(sheetIndex);
                        } catch (IllegalArgumentException e) {
                            // ignore
                        }
                    } else {
                        ExcelWriteException e = new ExcelWriteException("sheetName or sheetIndex can't be null");
                        e.setCode(ExcelWriteException.CODE_OF_SHEET_NAME_AND_INDEX_IS_EMPTY);
                        throw e;
                    }
                    if (sheet == null) {
                        ExcelWriteException e = new ExcelWriteException("sheet of '" + sheetName + "' is not exist.");
                        e.setCode(ExcelWriteException.CODE_OF_SHEET_NOT_EXIST);
                        throw e;
                    }
                } else {
                    if (sheetName != null) {
                        sheet = workbook.getSheet(sheetName);
                        if (sheet != null) {
                            if (sheetIndex != null && !sheetIndex.equals(workbook.getSheetIndex(sheet))) {
                                ExcelWriteException e = new ExcelWriteException("sheetName[" + sheetName
                                                                                + "] and sheetIndex[" + sheetIndex
                                                                                + "] not match.");
                                e.setCode(ExcelWriteException.CODE_OF_SHEET_NAME_AND_INDEX_NOT_MATCH);
                                throw e;
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
                        ExcelWriteException e = new ExcelWriteException("sheetName or sheetIndex can't be null.");
                        e.setCode(ExcelWriteException.CODE_OF_SHEET_NAME_AND_INDEX_IS_EMPTY);
                        throw e;
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
                // write head
                writeHead(useTemplate, sheet, sheetProcessor);
                // sheet
                ExcelProcessControllerImpl controller = new ExcelProcessControllerImpl();
                int writeRowIndex = sheetProcessor.getStartRowIndex();
                boolean isBreak = false;
                Map<Integer, InnerRow> cacheForTemplateRow = new HashMap<Integer, InnerRow>();

                List<?> dataList = sheetProcessor.getDataList(); //
                if (dataList != null && !dataList.isEmpty()) {
                    for (Object rowData : dataList) {
                        // proc row
                        Row row = sheet.getRow(writeRowIndex);
                        if (row == null) {
                            row = sheet.createRow(writeRowIndex);
                        }
                        InnerRow templateRow = getTemplateRow(cacheForTemplateRow, sheet, sheetProcessor, writeRowIndex);
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
                        //
                        try {
                            controller.reset();
                            if (sheetProcessor.getRowProcessor() != null) {
                                sheetProcessor.getRowProcessor().process(controller, context, rowData, row);
                            }
                            if (!controller.isDoSkip()) {
                                writeRow(context, templateRow, row, rowData, sheetProcessor);
                                writeRowIndex++;
                            }
                            if (controller.isDoBreak()) {
                                isBreak = true;
                                break;
                            }
                        } catch (RuntimeException e) {
                            if (e instanceof ExcelWriteException) {
                                ExcelWriteException ewe = (ExcelWriteException) e;
                                // ef.setColIndex(null); user may want to set this value,
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
                    }
                    if (isBreak) {
                        break;
                    }
                }
                if (sheetProcessor.getTemplateStartRowIndex() != null
                    && sheetProcessor.getTemplateEndRowIndex() != null) {
                    writeDataValidations(sheet, sheetProcessor);
                    writeStyleAfterFinish(useTemplate, sheet, sheetProcessor);
                }
            } catch (Throwable e) {
                if (e instanceof ExcelWriteException) {
                    ExcelWriteException e0 = (ExcelWriteException) e;
                    e0.setSheetName(sheetName);
                    e0.setSheetIndex(sheetIndex);
                    sheetProcessor.onException(context, e0);
                } else {
                    ExcelWriteException e0 = new ExcelWriteException(e);
                    e0.setSheetName(sheetName);
                    e0.setSheetIndex(sheetIndex);
                    e0.setCode(ExcelWriteException.CODE_OF_PROCESS_EXCEPTION);
                    sheetProcessor.onException(context, e0);
                }
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

    private static InnerRow getTemplateRow(Map<Integer, InnerRow> cache, Sheet sheet,
                                           ExcelWriteSheetProcessor<?> sheetProcessor, int rowIndex) {
        InnerRow cachedRow = cache.get(rowIndex);
        if (cachedRow != null || cache.containsKey(rowIndex)) {
            return cachedRow;
        }
        InnerRow templateRow = null;
        if (sheetProcessor.getTemplateStartRowIndex() != null && sheetProcessor.getTemplateEndRowIndex() != null) {
            if (rowIndex <= sheetProcessor.getTemplateEndRowIndex()) {
                return null;
            }
            int tempRowIndex = (rowIndex - sheetProcessor.getTemplateEndRowIndex() - 1)
                               % (sheetProcessor.getTemplateEndRowIndex() - sheetProcessor.getTemplateStartRowIndex() + 1)
                               + sheetProcessor.getTemplateStartRowIndex();
            Row tempRow = sheet.getRow(tempRowIndex);
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
        cache.put(rowIndex, templateRow);
        return templateRow;
    }

    private static void writeStyleAfterFinish(boolean useTemplate, Sheet sheet,
                                              ExcelWriteSheetProcessor<?> sheetProcessor) {
        if (useTemplate) {
            return;
        }
        ExcelWriteFieldMapping excelWriteFieldMapping = sheetProcessor.getFieldMapping();
        if (excelWriteFieldMapping == null) {
            return;
        }
        Map<String, Map<Integer, ExcelWriteFieldMappingAttribute>> mme = excelWriteFieldMapping.export();
        if (mme == null) {
            return;
        }
        for (Map.Entry<String, Map<Integer, ExcelWriteFieldMappingAttribute>> entry : mme.entrySet()) {
            Map<Integer, ExcelWriteFieldMappingAttribute> me = entry.getValue();
            for (Integer column : me.keySet()) {
                sheet.autoSizeColumn(column);
            }
        }
    }

    private static void writeDataValidations(Sheet sheet, ExcelWriteSheetProcessor sheetProcessor) {
        int templateRowStartIndex = sheetProcessor.getTemplateStartRowIndex();
        int templateRowEndIndex = sheetProcessor.getTemplateEndRowIndex();
        int step = templateRowEndIndex - templateRowStartIndex + 1;
        int rowStartIndex = sheetProcessor.getStartRowIndex();

        Set<Integer> configColIndexSet = new HashSet<Integer>();
        for (Entry<String, Map<Integer, ExcelWriteFieldMappingAttribute>> fieldIndexMapping : sheetProcessor.getFieldMapping().export().entrySet()) {
            if (fieldIndexMapping == null || fieldIndexMapping.getValue() == null) {
                continue;
            }
            for (Entry<Integer, ExcelWriteFieldMappingAttribute> indexProcessorMapping : fieldIndexMapping.getValue().entrySet()) {
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
                    if (templateRowEndIndex < cellRangeAddress.getFirstRow()
                        || templateRowStartIndex > cellRangeAddress.getLastRow()) {// specify row
                        continue;
                    }
                    for (Integer configColIndex : configColIndexSet) {
                        if (configColIndex < cellRangeAddress.getFirstColumn()
                            || configColIndex > cellRangeAddress.getLastColumn()) {// specify column
                            continue;
                        }
                        if (templateRowStartIndex == templateRowEndIndex) {
                            newCellRangeAddressList.addCellRangeAddress(rowStartIndex, configColIndex,
                                                                        sheet.getLastRowNum(), configColIndex);
                            validationContains = true;
                        } else {
                            int start = cellRangeAddress.getFirstRow() > templateRowStartIndex ? cellRangeAddress.getFirstRow() : templateRowStartIndex;
                            int end = cellRangeAddress.getLastRow() < templateRowEndIndex ? cellRangeAddress.getLastRow() : templateRowEndIndex;
                            long lastRow = sheet.getLastRowNum();
                            if (lastRow > end) {
                                long count = (lastRow - templateRowEndIndex) / step;
                                int i = templateRowEndIndex;
                                for (; i < count; i++) {
                                    newCellRangeAddressList.addCellRangeAddress(start + i * step, configColIndex,
                                                                                end + i * step, configColIndex);
                                    validationContains = true;
                                }
                                long _start = start + i * step;
                                if (_start <= lastRow) {
                                    long _end = end + i * step;
                                    _end = _end < lastRow ? _end : lastRow;
                                    newCellRangeAddressList.addCellRangeAddress((int) _start, configColIndex,
                                                                                (int) _end, configColIndex);
                                    validationContains = true;
                                }
                            }
                        }
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

    private static void writeRow(ExcelWriteContext context, InnerRow templateRow, Row row, Object rowData,
                                 ExcelWriteSheetProcessor sheetProcessor) {
        boolean useTemplate = false;
        if (templateRow != null) {
            useTemplate = true;
        }
        ExcelWriteFieldMapping fieldMapping = sheetProcessor.getFieldMapping();
        for (Entry<String, Map<Integer, ExcelWriteFieldMappingAttribute>> entry : fieldMapping.export().entrySet()) {
            String fieldName = entry.getKey();
            Map<Integer, ExcelWriteFieldMappingAttribute> map = entry.getValue();
            for (Map.Entry<Integer, ExcelWriteFieldMappingAttribute> fieldValueMapping : map.entrySet()) {
                Integer colIndex = fieldValueMapping.getKey();
                ExcelWriteFieldMappingAttribute attribute = fieldValueMapping.getValue();
                Object val = null;
                if (rowData != null) {
                    val = getFieldValue(rowData, fieldName, sheetProcessor.isTrimSpace());
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

                ExcelWriteCellValueMapping valueMapping = attribute.getValueMapping();
                ExcelWriteCellProcessor processor = attribute.getCellProcessor();
                if (valueMapping != null) {
                    String key = null;
                    if (val != null) {
                        key = val.toString();
                    }
                    Object cval = valueMapping.get(key);
                    if (cval != null) {
                        writeCell(row.getRowNum(), colIndex, cell, cval, useTemplate, attribute, rowData);
                    } else {
                        if (!valueMapping.containsKey(key)) {
                            if (valueMapping.isSettedDefaultValue()) {
                                if (valueMapping.isSettedDefaultValueWithDefaultInput()) {
                                    writeCell(row.getRowNum(), colIndex, cell, val, useTemplate, attribute, rowData);
                                } else {
                                    writeCell(row.getRowNum(), colIndex, cell, valueMapping.getDefaultValue(),
                                              useTemplate, attribute, rowData);
                                }
                            } else if (valueMapping.getDefaultProcessor() != null) {
                                valueMapping.getDefaultProcessor().process(context, rowData, cell);
                            } else {
                                ExcelWriteException ex = new ExcelWriteException("Field value is " + key);
                                ex.setCode(ExcelWriteException.CODE_OF_FIELD_VALUE_NOT_MATCH);
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
                    writeCell(cell, val, useTemplate, attribute, rowData);
                    try {
                        processor.process(context, val, cell);
                    } catch (Throwable e) {
                        if (e instanceof ExcelWriteException) {
                            ExcelWriteException e0 = (ExcelWriteException) e;
                            e0.setColIndex(colIndex);
                            e0.setRowIndex(row.getRowNum());
                            throw e0;
                        } else {
                            ExcelWriteException e0 = new ExcelWriteException(e);
                            e0.setRowIndex(row.getRowNum());
                            e0.setColIndex(colIndex);
                            e0.setCode(ExcelWriteException.CODE_OF_PROCESS_EXCEPTION);
                            throw e0;
                        }
                    }
                } else {
                    writeCell(cell, val, useTemplate, attribute, rowData);
                }
            }
        }
    }

    @SuppressWarnings("rawtypes")
    private static Object getFieldValue(Object obj, String fieldName, boolean isTrimSpace) {
        Object val = null;
        if (obj instanceof Map) {
            val = ((Map) obj).get(fieldName);
        } else {// java bean
            val = getProperty(obj, fieldName);
        }
        // trim
        if (val != null && val instanceof String && isTrimSpace) {
            val = ((String) val).trim();
            if ("".equals(val)) {
                val = null;
            }
        }
        return val;
    }

    private static void writeCell(int rowIndex, int colIndex, Cell cell, Object val, boolean userTemplate,
                                  ExcelWriteFieldMappingAttribute attribute, Object bean) {
        try {
            writeCell(cell, val, userTemplate, attribute, bean);
        } catch (Throwable e) {
            if (e instanceof ExcelWriteException) {
                ExcelWriteException e0 = new ExcelWriteException();
                e0.setRowIndex(rowIndex);
                e0.setColIndex(colIndex);
                throw e0;
            } else {
                ExcelWriteException e0 = new ExcelWriteException(e);
                e0.setRowIndex(rowIndex);
                e0.setColIndex(colIndex);
                e0.setCode(ExcelWriteException.CODE_OF_PROCESS_EXCEPTION);
                throw e0;
            }
        }
    }

    public static void writeCell(Cell cell, Object val) {
        if (cell.getCellStyle() != null && cell.getCellStyle().getDataFormat() > 0) {
            writeCell(cell, val, true, null, null);
        } else {
            writeCell(cell, val, false, null, null);
        }
    }

    @SuppressWarnings("unused")
    private static void writeCell(Cell cell, Object val, boolean userTemplate,
                                  ExcelWriteFieldMappingAttribute attribute, Object bean) {
        if (attribute != null && attribute.getLinkField() != null) {
            String addressFieldName = attribute.getLinkField();
            String address = null;
            if (bean != null) {
                address = (String) getFieldValue(bean, addressFieldName, true);
            }
            Workbook wb = cell.getRow().getSheet().getWorkbook();

            Hyperlink link = wb.getCreationHelper().createHyperlink(attribute.getLinkType());
            link.setAddress(address);
            cell.setHyperlink(link);
            // Its style can't inherit from cell.
            CellStyle style = wb.createCellStyle();
            Font hlinkFont = wb.createFont();
            hlinkFont.setUnderline(Font.U_SINGLE);
            hlinkFont.setColor(IndexedColors.BLUE.getIndex());
            style.setFont(hlinkFont);
            if (cell.getCellStyle() != null) {
                style.setFillBackgroundColor(cell.getCellStyle().getFillBackgroundColor());
            }
            cell.setCellStyle(style);
        }
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

    private static class ExcelProcessControllerImpl implements ExcelProcessController {

        private boolean doSkip  = false;
        private boolean doBreak = false;

        public boolean isDoSkip() {
            return doSkip;
        }

        public boolean isDoBreak() {
            return doBreak;
        }

        public void doSkip() {
            this.doSkip = true;
        }

        public void doBreak() {
            this.doBreak = true;
        }

        public void reset() {
            this.doBreak = false;
            this.doSkip = false;
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

    private static Object getProperty(Object obj, String fieldName) {
        PropertyDescriptor pd = getPropertyDescriptor(obj.getClass(), fieldName);
        if (pd == null || pd.getReadMethod() == null) {
            throw new IllegalStateException("In class" + obj.getClass() + ", no getter method found for field '"
                                            + fieldName + "'");
        }
        try {
            return pd.getReadMethod().invoke(obj, (Object[]) null);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    private static void setProperty(Object obj, String fieldName, Object value) throws IllegalAccessException,
                                                                               IllegalArgumentException,
                                                                               InvocationTargetException {
        PropertyDescriptor pd = getPropertyDescriptor(obj.getClass(), fieldName);
        if (pd == null || pd.getWriteMethod() == null) {
            throw new IllegalStateException("In class" + obj.getClass() + "no setter method found for field '"
                                            + fieldName + "'");
        }
        Class<?> paramType = pd.getWriteMethod().getParameterTypes()[0];
        if (value != null && !paramType.isAssignableFrom(value.getClass())) {
            value = TypeUtils.cast(value, paramType, null);
        }
        pd.getWriteMethod().invoke(obj, value);
    }

    private static PropertyDescriptor getPropertyDescriptor(Class<?> clazz, String propertyName) {
        return BeanUtils.getPropertyDescriptor(clazz, propertyName);
    }
}
