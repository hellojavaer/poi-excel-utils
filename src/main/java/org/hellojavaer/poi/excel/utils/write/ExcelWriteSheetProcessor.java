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

import org.hellojavaer.poi.excel.utils.common.Assert;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public abstract class ExcelWriteSheetProcessor<T> {

    public abstract void beforeProcess(ExcelWriteContext<T> context);

    public abstract List<T> getDataList(ExcelWriteContext<T> context);

    public abstract void onException(ExcelWriteContext<T> context, RuntimeException e);

    public abstract void afterProcess(ExcelWriteContext<T> context);

    private Integer                   sheetIndex;
    private String                    sheetName;
    private int                       rowStartIndex = 0;
    private Integer                   templateRowStartIndex;
    private Integer                   templateRowEndIndex;
    private ExcelWriteFieldMapping    fieldMapping;
    private ExcelWriteRowProcessor<T> rowProcessor;
    private boolean                   trimSpace     = false;

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }

    public String getSheetName() {
        return sheetName;
    }

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

    public void setTemplateRows(Integer start, Integer end) {
        Assert.isTrue(start != null && end != null && start <= end || start == null && end == null);
        this.templateRowStartIndex = start;
        this.templateRowEndIndex = end;
    }

    public Integer getTemplateRowStartIndex() {
        return templateRowStartIndex;
    }

    public Integer getTemplateRowEndIndex() {
        return templateRowEndIndex;
    }

    public ExcelWriteRowProcessor<T> getRowProcessor() {
        return rowProcessor;
    }

    public void setRowProcessor(ExcelWriteRowProcessor<T> rowProcessor) {
        this.rowProcessor = rowProcessor;
    }

    public boolean isTrimSpace() {
        return trimSpace;
    }

    public void setTrimSpace(boolean trimSpace) {
        this.trimSpace = trimSpace;
    }
}
