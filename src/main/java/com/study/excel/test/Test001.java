package com.study.excel.test;

import java.text.DecimalFormat;
import java.text.NumberFormat;

public class Test001 {

	public static void main(String[] args) {
		NumberFormat numberFormater  = new DecimalFormat("##");
		// InputStream input = ServletActionContext.getServletContext().getResourceAsStream("resources/template/asset/asset_origin_change_asset_export.xlsx");
		System.out.print(numberFormater.format(23423432.234324));

	}

}
