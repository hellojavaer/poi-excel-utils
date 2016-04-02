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

    private static final long                                         serialVersionUID = 1L;

    private Map<Integer, Map<String, ExcelReadFieldMappingAttribute>> fieldMapping     = new LinkedHashMap<Integer, Map<String, ExcelReadFieldMappingAttribute>>();

    public ExcelReadFieldMappingAttribute put(String colIndex, String fieldName) {
        return put(ExcelUtils.convertColCharIndexToIntIndex(colIndex), fieldName);
    }

    public ExcelReadFieldMappingAttribute put(int colIndex, String fieldName) {
        Map<String, ExcelReadFieldMappingAttribute> map = fieldMapping.get(colIndex);
        if (map == null) {
            synchronized (fieldMapping) {
                if (fieldMapping.get(colIndex) == null) {
                    map = new ConcurrentHashMap<String, ExcelReadFieldMappingAttribute>();
                    fieldMapping.put(colIndex, map);
                }
            }
        }
        ExcelReadFieldMappingAttribute attribute = new ExcelReadFieldMappingAttribute();
        map.put(fieldName, attribute);
        return attribute;
    }

    public boolean isEmpty() {
        return fieldMapping.isEmpty();
    }

    public Set<Entry<Integer, Map<String, ExcelReadFieldMappingAttribute>>> entrySet() {
        return fieldMapping.entrySet();
    }

    public class ExcelReadFieldMappingAttribute {

        private boolean                   required = false;
        private ExcelReadCellProcessor    cellProcessor;
        private ExcelReadCellValueMapping valueMapping;

        public ExcelReadFieldMappingAttribute setRequired(boolean required) {
            this.required = required;
            return this;
        }

        public ExcelReadFieldMappingAttribute setCellProcessor(ExcelReadCellProcessor cellProcessor) {
            this.cellProcessor = cellProcessor;
            return this;
        }

        public ExcelReadFieldMappingAttribute setValueMapping(ExcelReadCellValueMapping valueMapping) {
            this.valueMapping = valueMapping;
            return this;
        }

        public boolean isRequired() {
            return required;
        }

        public ExcelReadCellProcessor getCellProcessor() {
            return cellProcessor;
        }

        public ExcelReadCellValueMapping getValueMapping() {
            return valueMapping;
        }

    }
}
