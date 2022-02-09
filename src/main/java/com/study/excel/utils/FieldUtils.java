package com.study.excel.utils;


import org.apache.commons.collections4.CollectionUtils;
import org.apache.log4j.Logger;

import com.study.excel.annoation.ExportField;
import com.study.excel.convertor.Convertor;
import com.study.excel.dto.FieldInfo;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;

public class FieldUtils {

	public static final String SET_PREFIX = "set";
	public static final String GET_PREFIX = "get";
	public static final String IS_PREFIX = "is";

	public static final String IMPORT_IN = "in";

	public static final String IMPORT_OUT = "out";

	private static final Logger LOGGER = Logger.getLogger(FieldUtils.class);

	/**
	 * @param prefix
	 * @param fieldName
	 * @return
	 */
	public static String getMethodNameForField(final String prefix, final String fieldName) {
		return prefix + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
	}

	/**
	 * 根据属性名，获到该属性的Get方法名
	 *
	 * @param fieldName
	 * @return
	 */
	public static String getGetMethodName(final String fieldName) {
		return getMethodNameForField(GET_PREFIX, fieldName);
	}

	/**
	 * 根据属性名，获到该属性的Get方法名
	 *
	 * @param fieldName
	 * @return
	 */
	public static String getSetMethodName(final String fieldName) {
		return getMethodNameForField(SET_PREFIX, fieldName);
	}

	/**
	 * @param cls
	 * @param methodName
	 * @return
	 */
	public static Method getMethodByMethodName(Class<?> cls, String methodName) {
		try {
			return cls.getMethod(methodName);
		} catch (NoSuchMethodException e) {
			LOGGER.warn("method  " + methodName + " not  found ");
		}
		return null;
	}

	/**
	 * @param cls
	 * @param fieldName
	 * @return
	 */
	public static Method getGetMethodByFieldName(Class<?> cls, String fieldName) {
		return getMethodByMethodName(cls, getGetMethodName(fieldName));

	}

	public static Method getMethodByMethodNameAndParametrTypes(Class<?> cls, String methodName,
			Class<?>... parameterTypes) {
		try {
			return cls.getMethod(methodName, parameterTypes);
		} catch (NoSuchMethodException e) {
			LOGGER.error("  method " + methodName + " not  found ");
		}
		return null;

	}

	/**
	 * @param cls
	 * @param fieldName
	 * @return
	 */
	public static Method getSetMethodByFieldName(Class<?> cls, String fieldName, Class<?>... parameterTypes) {
		String setMethodName = getSetMethodName(fieldName);
		try {
			return cls.getMethod(setMethodName, parameterTypes);
		} catch (NoSuchMethodException e) {
			LOGGER.error(
					"get  field    " + fieldName + " ,setMethod  failed !  method " + setMethodName + " not  found ");
		}

		return null;

	}

	/**
	 * cls 必须有无参的构造函数
	 *
	 * @param cls
	 * @param cls
	 * @return
	 */
	public static <T> T getInstanceByClass(Class<T> cls) throws Exception {
		try {
			return cls.getDeclaredConstructor().newInstance();
		} catch (Exception e) {
			 LOGGER.error("get instance  of  " + cls.getClass() + "failed !");
			 throw  new  Exception(e);
		}
	
	}
	
	
	
	
	
	 /**
     * @param cls 获出需要导出实体类字段，及字段title
     * @return
     * @throws Exception
     */
    public static FieldInfo[] getExplortHeadFieldDtos(Class<?> cls) throws Exception {
        List<FieldInfo> fieldDtoList = FieldUtils.getFieldInfo(cls);
        if (CollectionUtils.isEmpty(fieldDtoList)) {
            throw new Exception("there  is  no  filed to export !");
        }
        //判断order是否有相同，若有相同，抛出异常
        List<Integer> orderList = new ArrayList<>();
        for (FieldInfo dto : fieldDtoList) {
        	if(dto.getOrder()==null ||  dto.getOrder() < 0){
				throw new Exception("field order   on field [" + dto.getPropertyName() + "]  must  not  null  and  big than  zero");
			}
			if (orderList.contains(dto.getOrder())) {
				throw new Exception("Repeated export field order   on field [" + dto.getPropertyName() + "]");
			}
			orderList.add(dto.getOrder());
        }
        //排序，找出order最大值
        Collections.sort(orderList);
        int size = orderList.size();
        int max = orderList.get(size - 1);
        if (max > 1000) {
            throw new Exception("the  order value [" + max + " ] is  big  max  value   [" + 1000 + "]");
        }
        int min = orderList.get(0);
        if (min < 1) {
            throw new Exception("the  order value start with 1  ");
        }
        FieldInfo[] filedHeadArray = new FieldInfo[max];
        //优先排序cols>0
        int  i=0;
		//排序
		Collections.sort(fieldDtoList);
        for (FieldInfo dto : fieldDtoList) {
            	filedHeadArray[i++] = dto;
        }
        return filedHeadArray;
    }

	
	
	
	
	
	
	
	
	

	/**
	 * 获取需要导入，导出的字段信息。
	 * 
	 * @param cls
	 * @return
	 * @throws Exception 
	 */
	private static  List<FieldInfo> getFieldInfo(Class<?> cls) throws Exception {
		List<FieldInfo> fieldInfoList = new ArrayList<FieldInfo>();
		Field[] fields = cls.getDeclaredFields();
		FieldInfo fieldInfo = null;
		for (Field field : fields) {
			ExportField exportAnnotation = field.getAnnotation(ExportField.class);
			if (null == exportAnnotation) {
				continue;
			}
			fieldInfo = new FieldInfo();
			setMethods(cls, field, fieldInfo);
			fieldInfo.setTitle(exportAnnotation.title());
			fieldInfo.setOrder(exportAnnotation.order());
			fieldInfo.setDateFormat(exportAnnotation.dateFormat());
			fieldInfo.setNumberFromat(exportAnnotation.numberFromat());
			fieldInfo.setWidth(exportAnnotation.width());
			Class<?>  convertorCls  = exportAnnotation.convertor();
			if(convertorCls != ExportField.DEFAULT.class) {
				fieldInfo.setConvertor((Convertor<?>) FieldUtils.getInstanceByClass(convertorCls));
			}
			fieldInfoList.add(fieldInfo);

		}
		return fieldInfoList;

	}

	/**
	 * 设置属性的 get, set 方法
	 * 
	 * @param cls
	 * @param field
	 * @param fieldInfo
	 * @throws Exception
	 */
	private static void setMethods(Class<?> cls, Field field, FieldInfo fieldInfo) throws Exception {
		String fieldName = field.getName();
		Method getMethod = FieldUtils.getGetMethodByFieldName(cls, fieldName);
		if (getMethod == null) {
			final String booleanGetterName = FieldUtils.getMethodNameForField(FieldUtils.IS_PREFIX, fieldName);
			getMethod = FieldUtils.getMethodByMethodName(cls, booleanGetterName);
		}
		if (getMethod == null) {
			getMethod = FieldUtils.getMethodByMethodName(cls, fieldName);
		}
		if (getMethod == null) {
			throw new Exception("the  field " + fieldName + "  not  have a getMethod!");
		}
		Method setMethod = FieldUtils.getSetMethodByFieldName(cls, fieldName, field.getType());
		if (setMethod == null) {
			setMethod = FieldUtils.getMethodByMethodNameAndParametrTypes(cls,
					FieldUtils.SET_PREFIX + "" + fieldName.substring(2, fieldName.length()), field.getType());
		}
		if (setMethod == null) {
			throw new Exception("the  field " + fieldName + "  not  have a setMethod!");
		}
		fieldInfo.setPropertyName(fieldName);
		fieldInfo.setReaderMethod(getMethod);
		fieldInfo.setWriteMethod(setMethod);
		fieldInfo.setFieldType(field.getType());
	}

	
    /**
     * @param headDtos 获出需要导出字段title数组
     * @return
     * @throws Exception
     */
    public static String[] getHeadFieldTitles(FieldInfo[] headDtos) {
        String[] headTitles = new String[headDtos.length];
        int i = 0;
        for (FieldInfo dto : headDtos) {
            if (dto == null) {
                headTitles[i] = " ";
            } else {
                headTitles[i] = dto.getTitle();
            }
            i++;
        }
        return headTitles;
    }
}
