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

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class InnerWriteCellProcessorWrapper implements Serializable {

    private static final long          serialVersionUID = 1L;

    @SuppressWarnings("rawtypes")
    private ExcelWriteCellProcessor    processor;
    private ExcelWriteCellValueMapping valueMapping;

    @SuppressWarnings("rawtypes")
    public InnerWriteCellProcessorWrapper(ExcelWriteCellValueMapping valueMapping, ExcelWriteCellProcessor processor) {
        // this.fieldName = fieldName;
        this.valueMapping = valueMapping;
        this.processor = processor;
    }

    @SuppressWarnings("rawtypes")
    public ExcelWriteCellProcessor getProcessor() {
        return processor;
    }

    @SuppressWarnings("rawtypes")
    public void setProcessor(ExcelWriteCellProcessor processor) {
        this.processor = processor;
    }

    public ExcelWriteCellValueMapping getValueMapping() {
        return valueMapping;
    }

    public void setValueMapping(ExcelWriteCellValueMapping valueMapping) {
        this.valueMapping = valueMapping;
    }

}
