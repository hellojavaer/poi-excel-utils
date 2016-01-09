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
package org.hellojavaer.poi.excel.utils.write;

import java.util.List;

import org.apache.poi.ss.usermodel.Row;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public abstract class ExcelWriteSheetProcessor<T> {

    public abstract void beforeProcess(ExcelWriteContext<T> context);

    public abstract List<T> getDataList(ExcelWriteContext<T> context);

    public abstract Row process(ExcelWriteContext<T> context, T t, Row row);

    public abstract void onException(ExcelWriteContext<T> context, RuntimeException e);

    public abstract void afterProcess(ExcelWriteContext<T> context);

    private Integer                sheetIndex;
    private String                 sheetName;
    private int                    rowStartIndex = 0;
    private Integer                templateRowIndex;
    private ExcelWriteFieldMapping fieldMapping;

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    /**
     * if you have setted( or will set) templateRowIndex, sheetIndex (or
     * SheetName) is required. <br>
     * else this parameter is unnecessary.
     * 
     * @see #setSheetName
     * @see #setTemplateRowIndex
     */
    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public String getSheetName() {
        return sheetName;
    }

    /**
     * if you have setted( or will set) templateRowIndex, sheetIndex (or
     * SheetName) is required. <br>
     * else this parameter is unnecessary.
     * 
     * @see #setSheetIndex
     * @see #setTemplateRowIndex
     */
    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public int getRowStartIndex() {
        return rowStartIndex;
    }

    public void setRowStartIndex(int rowStartIndex) {
        this.rowStartIndex = rowStartIndex;
    }

    public ExcelWriteFieldMapping getFieldMapping() {
        return fieldMapping;
    }

    public void setFieldMapping(ExcelWriteFieldMapping fieldMapping) {
        this.fieldMapping = fieldMapping;
    }

    /**
     * @return
     */
    public Integer getTemplateRowIndex() {
        return templateRowIndex;
    }

    /**
     * if set 'templateRowIndex' and
     * 
     * @param templateRowIndex
     */
    public void setTemplateRowIndex(Integer templateRowIndex) {
        this.templateRowIndex = templateRowIndex;
    }

}
