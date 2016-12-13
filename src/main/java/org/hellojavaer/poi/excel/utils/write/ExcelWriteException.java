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

    private static final long serialVersionUID                       = 1L;

    public static final int   CODE_OF_PROCESS_EXCEPTION              = 0;
    public static final int   CODE_OF_SHEET_NOT_EXIST                = 20;
    public static final int   CODE_OF_SHEET_NAME_AND_INDEX_NOT_MATCH = 21;
    public static final int   CODE_OF_SHEET_NAME_AND_INDEX_IS_EMPTY  = 22;
    public static final int   CODE_OF_FIELD_VALUE_NOT_MATCH          = 61;

    private String            sheetName                              = null;
    private Integer           sheetIndex                             = null;
    private Integer           rowIndex                               = null;
    private String            colStrIndex                            = null;
    private Integer           colIndex                               = null;

    /**
     * [0-99] are system reserved values.user-define value should be larger than
     * or equal to 100.
     */
    private int               code                                   = CODE_OF_PROCESS_EXCEPTION;

    private String            msg;

    public ExcelWriteException() {
        super();
    }

    public ExcelWriteException(String message, Throwable cause) {
        super(message, cause);
        this.msg = message;
    }

    public ExcelWriteException(String message) {
        super(message);
        this.msg = message;
    }

    public ExcelWriteException(Throwable cause) {
        super(cause);
    }

    @Override
    public String getMessage() {
        StringBuilder sb = new StringBuilder();
        sb.append(this.msg);
        sb.append(" Detail coordinate is sheet:");
        sb.append(this.getSheetName());
        sb.append("[index:");
        sb.append(this.getSheetIndex());
        sb.append("]");
        sb.append(" row:");
        if (this.getRowIndex() != null) {
            sb.append(this.getRowIndex() + 1);
        } else {
            sb.append("null");
        }
        sb.append(" column:");
        sb.append(this.getColStrIndex());
        sb.append("[index:");
        sb.append(this.getColIndex());
        sb.append("], code is ");
        sb.append(this.getCode());
        sb.append(", and description is '");
        switch (this.getCode()) {
            case CODE_OF_PROCESS_EXCEPTION:
                sb.append("process exception");
                break;
            case CODE_OF_SHEET_NOT_EXIST:
                sb.append("sheet not exist");
                break;
            case CODE_OF_FIELD_VALUE_NOT_MATCH:
                sb.append("field value not match");
                break;
        }
        sb.append("'.");
        return sb.toString();
    }

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
