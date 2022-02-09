package com.study.excel.convertor;

import org.apache.commons.lang3.StringUtils;

public class Y_OR_N_Convertor extends Convertor<String> {

    @Override
    public String inport(Object obj) {
        if (obj != null && StringUtils.isNotBlank((String) obj)) {
            String cellValue = (String) obj;
            if ("是".equals(cellValue)) {
                return "Y";
            }

        }
        return "N";
    }

    @Override
    public Object export(String t) {
        if (t != null) {
            if ("Y".equals(t)) {
                return "是";
            }
            return "否";
        }

        return "";
    }
}
