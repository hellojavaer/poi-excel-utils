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
import java.util.LinkedHashMap;
import java.util.Map;
import java.util.Map.Entry;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;

import org.hellojavaer.poi.excel.utils.ExcelUtils;

/**
 * Config the mapping between excel column(by index) to Object field(by name).
 * 
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ExcelWriteFieldMapping implements Serializable {

    private static final long                                         serialVersionUID = 1L;

    private Map<String, Map<Integer, InnerWriteCellProcessorWrapper>> fieldMapping     = new LinkedHashMap<String, Map<Integer, InnerWriteCellProcessorWrapper>>();

    public void put(int colIndex, String fieldName) {
        put(colIndex, fieldName, null, null);
    }

    public void put(int colIndex, String fieldName, @SuppressWarnings("rawtypes")
    ExcelWriteCellProcessor proccessor) {
        put(colIndex, fieldName, null, proccessor);
    }

    public void put(int colIndex, String fieldName, ExcelWriteCellValueMapping valueMapping) {
        put(colIndex, fieldName, valueMapping, null);
    }

    public void put(String colIndex, String fieldName) {
        put(colIndex, fieldName, null, null);
    }

    public void put(String colIndex, String fieldName, ExcelWriteCellValueMapping valueMapping) {
        put(colIndex, fieldName, valueMapping, null);
    }

    public void put(String colIndex, String fieldName, @SuppressWarnings("rawtypes")
    ExcelWriteCellProcessor proccessor) {
        put(colIndex, fieldName, null, proccessor);
    }

    private void put(int colIndex, String fieldName, ExcelWriteCellValueMapping valueMapping,
                     @SuppressWarnings("rawtypes")
                     ExcelWriteCellProcessor proccessor) {
        Map<Integer, InnerWriteCellProcessorWrapper> map = fieldMapping.get(fieldName);
        if (map == null) {
            synchronized (fieldMapping) {
                if (fieldMapping.get(colIndex) == null) {
                    map = new ConcurrentHashMap<Integer, InnerWriteCellProcessorWrapper>();
                    fieldMapping.put(fieldName, map);
                }
            }
        }
        map.put(colIndex, new InnerWriteCellProcessorWrapper(valueMapping, proccessor));
    }

    private void put(String colIndex, String fieldName, ExcelWriteCellValueMapping valueMapping,
                     @SuppressWarnings("rawtypes")
                     ExcelWriteCellProcessor proccessor) {
        put(ExcelUtils.convertColCharIndexToIntIndex(colIndex), fieldName, valueMapping, proccessor);
    }

    public boolean isEmpty() {
        return fieldMapping.isEmpty();
    }

    public Set<Entry<String, Map<Integer, InnerWriteCellProcessorWrapper>>> entrySet() {
        return fieldMapping.entrySet();
    }

}
