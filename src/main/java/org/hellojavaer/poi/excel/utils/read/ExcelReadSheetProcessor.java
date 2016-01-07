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
package org.hellojavaer.poi.excel.utils.read;

import java.util.List;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public abstract class ExcelReadSheetProcessor<T> {

    private Integer                  sheetIndex;
    private String                   sheetName;
    private Class<T>                 targetClass;
    private int                      rowStartIndex = 0;
    private Integer                  rowEndIndex;
    private Integer                  pageSize;
    private ExcelReadFieldMapping    fieldMapping;
    private ExcelReadRowProcessor<T> rowProcessor;

    public abstract void beforeProcess(ExcelReadContext<T> context);

    public abstract void process(ExcelReadContext<T> context, List<T> list);

    public abstract void onExcepton(ExcelReadContext<T> context, RuntimeException e);

    public abstract void afterProcess(ExcelReadContext<T> context);

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    /**
     * required.it can be replaced with {@code setSheetName};
     * 
     * @param sheetIndex
     * @see #setSheetName
     */
    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    /**
     * required.it can be replaced with {@code setSheetIndex};
     * 
     * @param sheetIndex
     * @see #setSheetIndex
     */
    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    /**
     * required
     * 
     * @param sheetIndex
     */
    public int getRowStartIndex() {
        return rowStartIndex;
    }

    /**
     * not necessary ,but always set
     * 
     * @param startRowIndex
     */
    public void setRowStartIndex(int rowStartIndex) {
        this.rowStartIndex = rowStartIndex;
    }

    public Integer getPageSize() {
        return pageSize;
    }

    /**
     * not necessary
     * 
     * @param pageSize
     */
    public void setPageSize(Integer pageSize) {
        this.pageSize = pageSize;
    }

    public ExcelReadFieldMapping getFieldMapping() {
        return fieldMapping;
    }

    public void setFieldMapping(ExcelReadFieldMapping fieldMapping) {
        this.fieldMapping = fieldMapping;
    }

    public ExcelReadRowProcessor<T> getRowProcessor() {
        return rowProcessor;
    }

    /**
     * not necessary
     * 
     * @param rowProcessor
     */
    public void setRowProcessor(ExcelReadRowProcessor<T> rowProcessor) {
        this.rowProcessor = rowProcessor;
    }

    public Class<T> getTargetClass() {
        return targetClass;
    }

    /**
     * required
     * 
     * @param targetClass
     */
    public void setTargetClass(Class<T> targetClass) {
        this.targetClass = targetClass;
    }

    public Integer getRowEndIndex() {
        return rowEndIndex;
    }

    /**
     * not necessary
     * 
     * @param rowEndIndex
     */
    public void setRowEndIndex(Integer rowEndIndex) {
        this.rowEndIndex = rowEndIndex;
    }

}
