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

import org.hellojavaer.poi.excel.utils.common.Assert;

import java.util.List;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public abstract class ExcelWriteSheetProcessor<T> {

    public abstract void beforeProcess(ExcelWriteContext<T> context);

    public abstract void onException(ExcelWriteContext<T> context, ExcelWriteException e);

    public abstract void afterProcess(ExcelWriteContext<T> context);

    private Integer                   sheetIndex;
    private String                    sheetName;
    private int                       startRowIndex = 0;
    private Integer                   templateStartRowIndex;
    private Integer                   templateEndRowIndex;
    private ExcelWriteFieldMapping    fieldMapping;
    private ExcelWriteRowProcessor<T> rowProcessor;
    private boolean                   trimSpace     = false;
    private Integer                   headRowIndex;
    private List<T>                   dataList;
    private Integer                   theme;

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

    public int getStartRowIndex() {
        return startRowIndex;
    }

    public void setStartRowIndex(int startRowIndex) {
        this.startRowIndex = startRowIndex;
    }

    public ExcelWriteFieldMapping getFieldMapping() {
        return fieldMapping;
    }

    public void setFieldMapping(ExcelWriteFieldMapping fieldMapping) {
        this.fieldMapping = fieldMapping;
    }

    public void setTemplateRows(Integer start, Integer end) {
        Assert.isTrue(start != null && end != null && start <= end || start == null && end == null);
        this.templateStartRowIndex = start;
        this.templateEndRowIndex = end;
    }

    public Integer getTemplateStartRowIndex() {
        return templateStartRowIndex;
    }

    public Integer getTemplateEndRowIndex() {
        return templateEndRowIndex;
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

    public Integer getHeadRowIndex() {
        return headRowIndex;
    }

    public void setHeadRowIndex(Integer headRowIndex) {
        this.headRowIndex = headRowIndex;
    }

    public List<T> getDataList() {
        return dataList;
    }

    public void setDataList(List<T> dataList) {
        this.dataList = dataList;
    }

    public Integer getTheme() {
        return theme;
    }

    public void setTheme(Integer theme) {
        this.theme = theme;
    }

}
