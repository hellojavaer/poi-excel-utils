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
import java.util.concurrent.ConcurrentHashMap;

import org.hellojavaer.poi.excel.utils.common.Assert;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class ExcelReadFieldMapping implements Serializable {

    private static final long                                        serialVersionUID = 1L;

    private Map<String, Map<String, ExcelReadFieldMappingAttribute>> fieldMapping     = new LinkedHashMap<String, Map<String, ExcelReadFieldMappingAttribute>>();

    public ExcelReadFieldMappingAttribute put(String colIndexOrColName, String fieldName) {
        Assert.notNull(colIndexOrColName);
        Assert.notNull(fieldName);
        Map<String, ExcelReadFieldMappingAttribute> map = fieldMapping.get(colIndexOrColName);
        if (map == null) {
            synchronized (fieldMapping) {
                if (fieldMapping.get(colIndexOrColName) == null) {
                    map = new ConcurrentHashMap<String, ExcelReadFieldMappingAttribute>();
                    fieldMapping.put(colIndexOrColName, map);
                }
            }
        }
        ExcelReadFieldMappingAttribute attribute = new ExcelReadFieldMappingAttribute();
        map.put(fieldName, attribute);
        return attribute;
    }

    public Map<String, Map<String, ExcelReadFieldMappingAttribute>> export() {
        return fieldMapping;
    }

    public class ExcelReadFieldMappingAttribute implements Serializable {

        private static final long         serialVersionUID = 1L;

        private boolean                   required         = false;
        private ExcelReadCellProcessor    cellProcessor;
        private ExcelReadCellValueMapping valueMapping;
        private String                    linkField;

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

        public ExcelReadFieldMappingAttribute setLinkField(String linkField) {
            this.linkField = linkField;
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

        public String getLinkField() {
            return linkField;
        }

    }
}
