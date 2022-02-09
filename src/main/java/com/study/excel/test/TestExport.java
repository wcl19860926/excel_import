package com.study.excel.test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.study.excel.utils.ExcelUtils;

public class TestExport {
	
	
	public static void  main(String[] args) {
		
		try {

			Workbook      	workbook = new XSSFWorkbook();
			FileOutputStream  out  = new  FileOutputStream(new File("d://dd.xlsx"));
			List<AllocationVo>   allocationVoList  = new  ArrayList<AllocationVo>();
			
			AllocationVo   vo  = new  AllocationVo();
			vo.setAssetNo("你好唉");
			vo.setAssetNumber("123123214");
			vo.setCostCenter("中力你发");
			vo.setIsEnvironment(Boolean.FALSE);
			vo.setRemark("中华人民共和国");
			vo.setStorePlace("叫外卖");
			allocationVoList.add(vo);
			
			AllocationVo   vo1  = new  AllocationVo();
			vo1.setAssetNo("你好唉11111111");
			vo1.setAssetNumber("123123214");
			vo1.setCostCenter("中力你发111111111");
			vo1.setIsEnvironment(Boolean.FALSE);
			vo1.setRemark("中华人民共和国111111");
			vo1.setStorePlace("叫外卖1111111");
			allocationVoList.add(vo1);
			
			
			
			AllocationVo   vo2  = new  AllocationVo();
			vo2.setAssetNo("你好唉111112222222222");
			vo2.setAssetNumber("12312321422222222");
			vo2.setCostCenter("中力你发2222222222");
			vo2.setIsEnvironment(Boolean.FALSE);
			vo2.setRemark("中华人民共和国2222222");
			vo2.setStorePlace("叫外卖222222222");
			allocationVoList.add(vo2);
			Sheet  sheet  = workbook.createSheet();
			ExcelUtils.writeDateToSheet(workbook  ,sheet ,allocationVoList , 4,true);
		
			workbook.write(out);
			out.flush();
			out.close();
			
			
		} catch (Exception e) {
			
			e.printStackTrace();
		}
		
		
		
		
	}

}
