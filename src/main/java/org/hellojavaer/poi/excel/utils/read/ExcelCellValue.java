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

import java.io.Serializable;
import java.math.BigDecimal;
import java.math.BigInteger;
import java.util.Date;

import com.alibaba.fastjson.util.TypeUtils;

/**
 * provide rich methods for type castting.
 * 
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ExcelCellValue implements Serializable {

    private static final long serialVersionUID = 1L;
    private Object            originalValue;

    public ExcelCellValue(Object originalValue) {
        this.originalValue = originalValue;
    }

    public Byte getByteValue() {
        return TypeUtils.castToByte(originalValue);
    }

    public Short getShortValue() {
        return TypeUtils.castToShort(originalValue);
    }

    public Integer getIntValue() {
        return TypeUtils.castToInt(originalValue);
    }

    public Long getLongValue() {
        return TypeUtils.castToLong(originalValue);
    }

    public Float getFloatValue() {
        return TypeUtils.castToFloat(originalValue);
    }

    public Double getDoubleValue() {
        return TypeUtils.castToDouble(originalValue);
    }

    public String getStringValue() {
        return TypeUtils.castToString(originalValue);
    }

    public Boolean getBooleanValue() {
        return TypeUtils.castToBoolean(originalValue);
    }

    public Date getDateValue() {
        return TypeUtils.castToDate(originalValue);
    }

    public BigDecimal getBigDecimal() {
        return TypeUtils.castToBigDecimal(originalValue);
    }

    public BigInteger getBigInteger() {
        return TypeUtils.castToBigInteger(originalValue);
    }

    public java.sql.Timestamp getTimestamp() {
        return TypeUtils.castToTimestamp(originalValue);
    }

    public Object getOriginalValue() {
        return originalValue;
    }

    @Override
    public String toString() {
        return TypeUtils.castToString(originalValue);
    }
}
