/*
 * Copyright 2015-2015 the original author or authors.
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

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class InnerReadCellProcessorWrapper implements Serializable {

	private static final long serialVersionUID = 1L;

	private boolean required = true;
	private ExcelReadCellProcessor processor;
	private ExcelReadCellValueMapping valueMapping;

	public InnerReadCellProcessorWrapper(
			ExcelReadCellValueMapping valueMapping,
			ExcelReadCellProcessor processor, boolean required) {
		this.valueMapping = valueMapping;
		this.processor = processor;
		this.required = required;
	}

	public ExcelReadCellProcessor getProcessor() {
		return processor;
	}

	public void setProcessor(ExcelReadCellProcessor processor) {
		this.processor = processor;
	}

	public ExcelReadCellValueMapping getValueMapping() {
		return valueMapping;
	}

	public void setValueMapping(ExcelReadCellValueMapping valueMapping) {
		this.valueMapping = valueMapping;
	}

	public boolean isRequired() {
		return required;
	}

	public void setRequired(boolean required) {
		this.required = required;
	}

}
