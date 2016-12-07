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

import org.hellojavaer.poi.excel.utils.ExcelUtils;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ExcelWriteException extends RuntimeException {

    private static final long serialVersionUID                = 1L;

    public static final int   CODE_OF_SHEET_NOT_EXSIT         = 0;
    public static final int   CODE_OF_PROCESS_EXCEPTION       = 1;
    public static final int   CODE_OF_FIELD_VALUE_NOT_MATCHED = 2;

    private String            sheetName;
    private Integer           sheetIndex;
    private Integer           rowIndex                        = null;
    private String            colStrIndex                     = null;
    private Integer           colIndex                        = null;

    public String getSheetName() {
        return sheetName;
    }

    public void setSheetName(String sheetName) {
        this.sheetName = sheetName;
    }

    public Integer getSheetIndex() {
        return sheetIndex;
    }

    public void setSheetIndex(Integer sheetIndex) {
        this.sheetIndex = sheetIndex;
    }


    /**
     * [0-99] are system reserved values.user-define value should be larger than
     * or equal to 100.
     */
    private int               code                            = CODE_OF_PROCESS_EXCEPTION;

    /**
     * [0-99] are system reserved values.user-define value should be larger than
     * or equal to 100.
     */
    public int getCode() {
        return code;
    }

    /**
     * [0-99] are system reserved values.user-define value should be larger than
     * or equal to 100.
     */
    public void setCode(int code) {
        this.code = code;
    }

    public ExcelWriteException() {
        super();
    }

    public ExcelWriteException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelWriteException(String message) {
        super(message);
    }

    public ExcelWriteException(Throwable cause) {
        super(cause);
    }

    public Integer getRowIndex() {
        return rowIndex;
    }

    public void setRowIndex(Integer rowIndex) {
        this.rowIndex = rowIndex;
    }

    public String getColStrIndex() {
        return colStrIndex;
    }

    public Integer getColIndex() {
        return colIndex;
    }

    public void setColStrIndex(String colStrIndex) {
        this.colStrIndex = colStrIndex;
        if (colStrIndex == null) {
            this.colIndex = null;
        } else {
            this.colIndex = ExcelUtils.convertColCharIndexToIntIndex(colStrIndex);
        }
    }
    
    public void setColIndex(Integer colIndex) {
        this.colIndex = colIndex;
        if (colIndex == null) {
            this.colStrIndex = null;
        } else {
            this.colStrIndex = ExcelUtils.convertColIntIndexToCharIndex(colIndex);
        }
    }

}
