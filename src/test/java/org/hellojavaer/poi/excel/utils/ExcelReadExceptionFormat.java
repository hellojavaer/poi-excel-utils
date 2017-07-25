package org.hellojavaer.poi.excel.utils;

import org.hellojavaer.poi.excel.utils.read.ExcelReadException;

/**
 *
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>,created on 08/12/2016.
 */
public class ExcelReadExceptionFormat {

    public static String format(ExcelReadException e) {
        StringBuilder sb = new StringBuilder();
        if (e.getCode() == ExcelReadException.CODE_OF_SHEET_NOT_EXIST) {
            sb.append("表:");
            if (e.getSheetName() != null) {
                sb.append(e.getSheetName());
            }
            if (e.getSheetIndex() != null) {
                sb.append('[');
                sb.append(e.getSheetIndex());
                sb.append(']');
            }
            sb.append("不存在");
            return sb.toString();
        }
        if (e.getCode() == ExcelReadException.CODE_OF_SHEET_NAME_AND_INDEX_NOT_MATCH) {
            sb.append("表名:");
            sb.append(e.getSheetName());
            sb.append("和表索引:[");
            sb.append(e.getSheetIndex());
            sb.append("]");
            sb.append("不匹配");
            return sb.toString();
        }
        if (e.getCode() == ExcelReadException.CODE_OF_SHEET_NAME_AND_INDEX_IS_EMPTY) {
            sb.append("表名和索引不能都为空");
            return sb.toString();
        }
        sb.append("表 ");
        sb.append(e.getSheetName());
        //sb.append("[");
        //sb.append(e.getSheetIndex());
        //sb.append("]");
        if (e.getRowIndex() == null) {
            sb.append(" 在处理表数据时异常");
        } else {
            sb.append(" , 第 ");
            sb.append(e.getRowIndex());
            sb.append(" 行");
            if (e.getColIndex() == null) {
                sb.append("时处理异常");
            } else {
                sb.append(", 第 ");
                sb.append(e.getColStrIndex());
                sb.append(" 列");
                if (e.getCode() == ExcelReadException.CODE_OF_PROCESS_EXCEPTION) {
                    sb.append(",数据处理异常,请校验数据格式的正确性.");
                } else if (e.getCode() == ExcelReadException.CODE_OF_CELL_VALUE_NOT_MATCH) {
                    sb.append(",所输入的值不合法,请校验数据格式的正确性.");
                } else if (e.getCode() == ExcelReadException.CODE_OF_CELL_VALUE_REQUIRED) {
                    sb.append("是必填项.");
                }
            }
        }
        return sb.toString();
    }

}
