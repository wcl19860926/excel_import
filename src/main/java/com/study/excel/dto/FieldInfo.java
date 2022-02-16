package com.study.excel.dto;


import java.lang.reflect.Method;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.CellStyle;

import com.study.excel.convertor.Convertor;
import com.study.excel.utils.ExcelUtils;



public class FieldInfo   implements   Comparable<FieldInfo>{
	
	/**
	 * excel 导入的表头行名子
	 */
	private  String  title ;
	
	
	/**
	 * 属性名
	 */
	private  String  propertyName;
	
	/**
	 * 解析字段的顺序
	 */
	private  Integer  order;
	
     /**
      *  导出到exlce的宽度
      */
	 private  Integer  width;
	
	/**
	 * 属性的读方法
	 */
	
	private  Method  readerMethod;
	/**
	 * 属性的写方法
	 */
	
	private  Method  writeMethod;
	
	
	/**
	 * 字段类型
	 */
	private  Class<?>  fieldType;
	
	
	/**
	 * 日期格式化时的格试
	 */
    private  String  dateFormat;
    
    /**
     * 数据格试
     */
    private  String  numberFromat;
    
    
    
    private  SimpleDateFormat  dataFormater ;
    
    private NumberFormat   numberFormater;
    
    
    private Convertor<?> convertor;


    CellStyle  cellStyle;

    private  Integer  alignment;




	public NumberFormat getNumberFormater() {
		return numberFormater;
	}



	public Convertor getConvertor() {
		return convertor;
	}





	public void setConvertor(Convertor convertor) {
		this.convertor = convertor;
	}





	public SimpleDateFormat getDataFormater() {
		return dataFormater;
	}





	public String getPropertyName() {
		return propertyName;
	}

	public void setPropertyName(String propertyName) {
		this.propertyName = propertyName;
	}

	public Integer getOrder() {
		return order;
	}

	public void setOrder(Integer order) {
		this.order = order;
	}

	public Method getReaderMethod() {
		return readerMethod;
	}

	public void setReaderMethod(Method readerMethod) {
		this.readerMethod = readerMethod;
	}

	public Method getWriteMethod() {
		return writeMethod;
	}

	public void setWriteMethod(Method writeMethod) {
		this.writeMethod = writeMethod;
	}

	public Class<?> getFieldType() {
		return fieldType;
	}

	public void setFieldType(Class<?> fieldType) {
		this.fieldType = fieldType;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
		if(StringUtils.isNotEmpty(dateFormat)) {
			dataFormater  = new SimpleDateFormat(dateFormat);
		}
	}

	public String getNumberFromat() {
		return numberFromat;
	}

	public void setNumberFromat(String numberFromat) {
		this.numberFromat = numberFromat;
		if(StringUtils.isNotEmpty(this.numberFromat)) {
			numberFormater  = new DecimalFormat(numberFromat);
		}
	}

	public String getTitle() {
		return title;
	}

	public void setTitle(String title) {
		this.title = title;
	}

	public Integer getWidth() {
		return width;
	}

	public void setWidth(Integer width) {
		this.width = width;
	}


	@Override
	public int compareTo(FieldInfo o) {
		if(o==null){
			return  0;
		}
		return this.order.compareTo(o.order);
	}


	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

	public Integer getAlignment() {
		return alignment;
	}

	public void setAlignment(Integer alignment) {
		this.alignment = alignment;
	}
}
