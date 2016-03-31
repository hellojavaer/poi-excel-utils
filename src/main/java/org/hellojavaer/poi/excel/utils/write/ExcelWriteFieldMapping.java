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

    private static final long                                          serialVersionUID = 1L;

    private Map<String, Map<Integer, ExcelWriteFieldMappingAttribute>> fieldMapping     = new LinkedHashMap<String, Map<Integer, ExcelWriteFieldMappingAttribute>>();

    public ExcelWriteFieldMappingAttribute put(String colIndex, String fieldName) {
        return put(ExcelUtils.convertColCharIndexToIntIndex(colIndex), fieldName);
    }

    public ExcelWriteFieldMappingAttribute put(int colIndex, String fieldName) {
        Map<Integer, ExcelWriteFieldMappingAttribute> map = fieldMapping.get(fieldName);
        if (map == null) {
            synchronized (fieldMapping) {
                if (fieldMapping.get(colIndex) == null) {
                    map = new ConcurrentHashMap<Integer, ExcelWriteFieldMappingAttribute>();
                    fieldMapping.put(fieldName, map);
                }
            }
        }
        ExcelWriteFieldMappingAttribute attribute = new ExcelWriteFieldMappingAttribute();
        map.put(colIndex, attribute);
        return attribute;
    }

    public boolean isEmpty() {
        return fieldMapping.isEmpty();
    }

    public Set<Entry<String, Map<Integer, ExcelWriteFieldMappingAttribute>>> entrySet() {
        return fieldMapping.entrySet();
    }

    public class ExcelWriteFieldMappingAttribute {

        @SuppressWarnings("rawtypes")
        private ExcelWriteCellProcessor    cellProcessor;
        private ExcelWriteCellValueMapping valueMapping;
        private String                     head;

        public ExcelWriteFieldMappingAttribute setCellProcessor(ExcelWriteCellProcessor cellProcessor) {
            this.cellProcessor = cellProcessor;
            return this;
        }

        public ExcelWriteFieldMappingAttribute setValueMapping(ExcelWriteCellValueMapping valueMapping) {
            this.valueMapping = valueMapping;
            return this;
        }

        public ExcelWriteFieldMappingAttribute setHead(String head) {
            this.head = head;
            return this;
        }

        public ExcelWriteCellProcessor getCellProcessor() {
            return cellProcessor;
        }

        public ExcelWriteCellValueMapping getValueMapping() {
            return valueMapping;
        }

        public String getHead() {
            return head;
        }

    }
}
