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

import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.hellojavaer.poi.excel.utils.ExcelUtils;
import org.hellojavaer.poi.excel.utils.read.ExcelCellValue;

/**
 * Writting excel context
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ExcelWriteContext<T> implements Serializable {

    private static final long   serialVersionUID = 1L;

    private Map<String, Object> contextContainer = new HashMap<String, Object>();

    private Sheet               curSheet;
    private Integer             curSheetIndex;
    private String              curSheetName;

    private Row                 curRow;
    private Integer             curRowIndex;

    private Cell                curCell;
    private Integer             curColIndex;
    private String              curColStrIndex;

    public Cell getCell(Integer rowIndex, Integer colIndex) {
        if (rowIndex == null || rowIndex < 0 || colIndex == null || colIndex < 0) {
            return null;
        }
        if (this.curSheet == null) {
            return null;
        } else {
            Row row = curSheet.getRow(rowIndex);
            if (row == null) {
                return null;
            } else {
                return row.getCell(colIndex);
            }
        }
    }

    public Cell getCell(Integer rowIndex, String colStrIndex) {
        return getCell(rowIndex, ExcelUtils.convertColCharIndexToIntIndex(colStrIndex));
    }

    public ExcelCellValue getCellValue(Integer rowIndex, Integer colIndex) {
        Cell cell = getCell(rowIndex, colIndex);
        return ExcelUtils.readCell(cell);
    }

    public ExcelCellValue getCellValue(Integer rowIndex, String colStrIndex) {
        return getCellValue(rowIndex, ExcelUtils.convertColCharIndexToIntIndex(colStrIndex));
    }

    public Cell setCellValue(Integer rowIndex, String colStrIndex, Object val) {
        return setCellValue(rowIndex, ExcelUtils.convertColCharIndexToIntIndex(colStrIndex), val);
    }

    // CellRangeAddress
    public Cell setCellValue(int firstRow, int lastRow, int firstCol, int lastCol, Object val) {
        CellRangeAddress cellRangeAddress = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
        curSheet.addMergedRegion(cellRangeAddress);
        return setCellValue(firstRow, firstCol, val);
    }

    public Cell setCellValue(int rowIndex, int colIndex, Object val) {
        if (rowIndex < 0 || colIndex < 0) {
            throw new IllegalArgumentException("rowIndex:" + rowIndex + " &colIndex:" + colIndex);
        }
        if (curSheet == null) {
            throw new NullPointerException("sheet is null");
        }
        Row row = curSheet.getRow(rowIndex);
        if (row == null) {
            row = curSheet.createRow(rowIndex);
        }
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        ExcelUtils.writeCell(cell, val);
        return cell;
    }

    public Sheet getCurSheet() {
        return curSheet;
    }

    public void setCurSheet(Sheet curSheet) {
        this.curSheet = curSheet;
    }

    public Integer getCurSheetIndex() {
        return curSheetIndex;
    }

    public void setCurSheetIndex(Integer curSheetIndex) {
        this.curSheetIndex = curSheetIndex;
    }

    public String getCurSheetName() {
        return curSheetName;
    }

    public void setCurSheetName(String curSheetName) {
        this.curSheetName = curSheetName;
    }

    public Row getCurRow() {
        return curRow;
    }

    public void setCurRow(Row curRow) {
        this.curRow = curRow;
    }

    public Integer getCurRowIndex() {
        return curRowIndex;
    }

    public void setCurRowIndex(Integer curRowIndex) {
        this.curRowIndex = curRowIndex;
    }

    public Cell getCurCell() {
        return curCell;
    }

    public void setCurCell(Cell curCell) {
        this.curCell = curCell;
    }

    public Integer getCurColIndex() {
        return curColIndex;
    }

    public void setCurColIndex(Integer curColIndex) {
        this.curColIndex = curColIndex;
        if (curColIndex == null) {
            this.curColStrIndex = null;
        } else {
            this.curColStrIndex = ExcelUtils.convertColIntIndexToCharIndex(curColIndex);
        }
    }

    public String getCurColStrIndex() {
        return curColStrIndex;
    }

    public void setAttribute(String key, Object value) {
        this.contextContainer.put(key, value);
    }

    public Object getAttribute(String key) {
        return this.contextContainer.get(key);
    }

}
