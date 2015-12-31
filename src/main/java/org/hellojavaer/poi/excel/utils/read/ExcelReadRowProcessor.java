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

import org.apache.poi.ss.usermodel.Row;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public interface ExcelReadRowProcessor<T> {

	/**
	 * if you return null,this row will be skipped.
	 */
	T process(ExcelReadContext<T> context, Row row, T t);
}
