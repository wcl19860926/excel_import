package com.study.excel.convertor;

public abstract  class Convertor<T> {

	public abstract T inport(Object obj);

	public abstract Object export(T t);

}
