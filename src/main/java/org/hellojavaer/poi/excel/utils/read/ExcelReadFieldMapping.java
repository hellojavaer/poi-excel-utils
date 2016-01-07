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
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

import org.hellojavaer.poi.excel.utils.ExcelUtils;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ExcelReadFieldMapping implements Serializable {

    private static final boolean                                     DEFAULT_REQUIRE  = true;
    private static final long                                        serialVersionUID = 1L;

    private Map<Integer, Map<String, InnerReadCellProcessorWrapper>> fieldMapping     = new LinkedHashMap<Integer, Map<String, InnerReadCellProcessorWrapper>>();

    public void put(int colIndex, String fieldName) {
        put(colIndex, fieldName, null, null, DEFAULT_REQUIRE);
    }

    public void put(int colIndex, String fieldName, boolean required) {
        put(colIndex, fieldName, null, null, required);
    }

    public void put(int colIndex, String fieldName, ExcelReadCellProcessor proccessor) {
        put(colIndex, fieldName, null, proccessor, DEFAULT_REQUIRE);
    }

    public void put(int colIndex, String fieldName, ExcelReadCellValueMapping valueMapping) {
        put(colIndex, fieldName, valueMapping, null, DEFAULT_REQUIRE);
    }

    public void put(int colIndex, String fieldName, ExcelReadCellProcessor proccessor, boolean required) {
        put(colIndex, fieldName, null, proccessor, required);
    }

    public void put(int colIndex, String fieldName, ExcelReadCellValueMapping valueMapping, boolean required) {
        put(colIndex, fieldName, valueMapping, null, required);
    }

    public void put(String colIndex, String fieldName) {
        put(colIndex, fieldName, null, null, DEFAULT_REQUIRE);
    }

    public void put(String colIndex, String fieldName, boolean required) {
        put(colIndex, fieldName, null, null, required);
    }

    public void put(String colIndex, String fieldName, ExcelReadCellValueMapping valueMapping) {
        put(colIndex, fieldName, valueMapping, null, DEFAULT_REQUIRE);
    }

    public void put(String colIndex, String fieldName, ExcelReadCellProcessor proccessor) {
        put(colIndex, fieldName, null, proccessor, DEFAULT_REQUIRE);
    }

    public void put(String colIndex, String fieldName, ExcelReadCellValueMapping valueMapping, boolean required) {
        put(colIndex, fieldName, valueMapping, null, required);
    }

    public void put(String colIndex, String fieldName, ExcelReadCellProcessor proccessor, boolean required) {
        put(colIndex, fieldName, null, proccessor, required);
    }

    private void put(int colIndex, String fieldName, ExcelReadCellValueMapping valueMapping,
                     ExcelReadCellProcessor proccessor, boolean required) {
        Map<String, InnerReadCellProcessorWrapper> map = fieldMapping.get(colIndex);
        if (map == null) {
            synchronized (fieldMapping) {
                if (fieldMapping.get(colIndex) == null) {
                    map = new ConcurrentHashMap<String, InnerReadCellProcessorWrapper>();
                    fieldMapping.put(colIndex, map);
                }
            }
        }
        map.put(fieldName, new InnerReadCellProcessorWrapper(valueMapping, proccessor, required));
    }

    private void put(String colIndex, String fieldName, ExcelReadCellValueMapping valueMapping,
                     ExcelReadCellProcessor proccessor, boolean required) {
        put(ExcelUtils.convertColCharIndexToIntIndex(colIndex), fieldName, valueMapping, proccessor, required);
    }

    public boolean isEmpty() {
        return fieldMapping.isEmpty();
    }

    public Set<Entry<Integer, Map<String, InnerReadCellProcessorWrapper>>> entrySet() {
        return fieldMapping.entrySet();
    }

}
