package com.study.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

import org.apache.poi.ss.usermodel.Workbook;

import com.study.excel.utils.ExcelUtils;

public class TestInport {
	
	
	public static void  main(String[] args) {
		
		try {
			InputStream  input = TestInport.class.getResourceAsStream("../utils/aa.xlsx");
			Workbook    workbook = 	ExcelUtils.createXssfWorkbook(input);
			List<AllocationVo>  resultDate  = ExcelUtils.readDateFromSheet(workbook.getSheetAt(0), AllocationVo.class, 3);
			System.out.print(resultDate);
			
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
		
		
	}

}
