package com.study.excel.annoation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import com.study.excel.convertor.Convertor;

@Target({ElementType.METHOD, ElementType.FIELD, ElementType.ANNOTATION_TYPE})
@Retention(RetentionPolicy.RUNTIME)
public @interface ExportField {

	  String  title()  default "";

	  /**
	   * 导入导出的顺序
	   * @return
	   */
	  int  order();
	  
	  /**
	   * 用于控制写入exlcel 单无格的宽度
	   * @return
	   */
	  int  width()  default  3;
	  

	  /**
	   *  用于导入导出日期格试
	   * @return
	   */
	  String  dateFormat()   default  "yyyy-MM-dd HH:mm:ss";
	  
	  
	  /**
	   * 用于控制导入导出，的数字格式
	   * @return
	   */
	  String  numberFromat()   default  "";
	  
	  /**
	   * 
	   * @return
	   */
	  
      Class<?  extends  Convertor<?>>  convertor()   default  DEFAULT.class;

      
      
      
      
      /**
       * 
       * 默认excel字段值转换
       *
       */
       static  class DEFAULT extends Convertor<Object> {

		@Override
		public Object inport(Object obj) {
			
			return obj;
		}

		@Override
		public Object export(Object obj) {
			
			return obj;
		}

    	

    	}
	
}
