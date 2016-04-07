package org.hellojavaer.poi.excel.utils;

import java.io.Serializable;
import java.util.Date;

/**
 * @author <a href="mailto:hellojavaer@gmail.com">zoukaiming</a>
 */
public class TestBean implements Serializable {

    private static final long serialVersionUID = 1L;

    private Byte              byteField;
    private Short             shortField;
    private Integer           intField;
    private Long              longField;
    private Float             floatField;
    private Double            doubleField;
    private Boolean           boolField;
    private String            stringField;
    private Date              dateField;
    private String            enumField1;
    private String            enumField2;
    private String            url;

    public Byte getByteField() {
        return byteField;
    }

    public void setByteField(Byte byteField) {
        this.byteField = byteField;
    }

    public Short getShortField() {
        return shortField;
    }

    public void setShortField(Short shortField) {
        this.shortField = shortField;
    }

    public Integer getIntField() {
        return intField;
    }

    public void setIntField(Integer intField) {
        this.intField = intField;
    }

    public Long getLongField() {
        return longField;
    }

    public void setLongField(Long longField) {
        this.longField = longField;
    }

    public Float getFloatField() {
        return floatField;
    }

    public void setFloatField(Float floatField) {
        this.floatField = floatField;
    }

    public Double getDoubleField() {
        return doubleField;
    }

    public void setDoubleField(Double doubleField) {
        this.doubleField = doubleField;
    }

    public Boolean getBoolField() {
        return boolField;
    }

    public void setBoolField(Boolean boolField) {
        this.boolField = boolField;
    }

    public String getStringField() {
        return stringField;
    }

    public void setStringField(String stringField) {
        this.stringField = stringField;
    }

    public Date getDateField() {
        return dateField;
    }

    public void setDateField(Date dateField) {
        this.dateField = dateField;
    }

    public String getEnumField1() {
        return enumField1;
    }

    public void setEnumField1(String enumField1) {
        this.enumField1 = enumField1;
    }

    public String getEnumField2() {
        return enumField2;
    }

    public void setEnumField2(String enumField2) {
        this.enumField2 = enumField2;
    }

    public String getUrl() {
        return url;
    }

    public void setUrl(String url) {
        this.url = url;
    }

}
