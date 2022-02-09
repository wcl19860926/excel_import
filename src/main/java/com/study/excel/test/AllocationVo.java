package com.study.excel.test;

import org.apache.commons.lang3.StringUtils;

import com.study.excel.annoation.ExportField;
import com.study.excel.convertor.Convertor;

public class AllocationVo {
	
	@ExportField(title="资产编码" ,order =1, width=1 )
	private  String  assetNo ;
	
	@ExportField(title="标签号" ,order =2, width=2 )
	private  String  assetNumber;
	
	@ExportField(title="成本中心" ,order =3, width=3 )
	private  String  costCenter;
	
	@ExportField(title="存放地点" ,order =4, width=3 )
	private  String  storePlace;
	
	@ExportField(title="是否环保设备" ,order =5, width=1, convertor=isEnvironmentConvertor.class)
	private  Boolean  isEnvironment;
	
	@ExportField(title="资产备注信息" ,order =6, width=3 )
	private  String   remark;

	public String getAssetNo() {
		return assetNo;
	}

	public void setAssetNo(String assetNo) {
		this.assetNo = assetNo;
	}

	public String getAssetNumber() {
		return assetNumber;
	}

	public void setAssetNumber(String assetNumber) {
		this.assetNumber = assetNumber;
	}

	public String getCostCenter() {
		return costCenter;
	}

	public void setCostCenter(String costCenter) {
		this.costCenter = costCenter;
	}

	public String getStorePlace() {
		return storePlace;
	}

	public void setStorePlace(String storePlace) {
		this.storePlace = storePlace;
	}

	public Boolean getIsEnvironment() {
		return isEnvironment;
	}

	public void setIsEnvironment(Boolean isEnvironment) {
		this.isEnvironment = isEnvironment;
	}

	public String getRemark() {
		return remark;
	}

	public void setRemark(String remark) {
		this.remark = remark;
	}
	
	
	
	
	public  static  class   isEnvironmentConvertor   extends  Convertor<Boolean>{

		@Override
		public Boolean inport(Object obj) {
		    if(obj!=null && StringUtils.isNotBlank((String)obj)) {
		    	String  cellValue  =  (String)obj;
		    	if("是".equals(cellValue)) {
		    		return  Boolean.TRUE;
		    	}
		    	
		    }
		    return    Boolean.FALSE;
		}

		@Override
		public Object export(Boolean t) {
			if(t!=null) {
				if(t) {
					return "是";
				}
				return  "否";
			}
			
			return "";
		}
		
		
		
	}
	
	
	
	

}
