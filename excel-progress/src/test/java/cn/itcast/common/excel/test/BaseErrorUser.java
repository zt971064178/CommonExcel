package cn.itcast.common.excel.test;

import cn.itcast.common.excel.annotation.ExcelColumn;
import cn.itcast.common.excel.annotation.ExcelHeader;
import cn.itcast.common.excel.annotation.ExcelWarning;
import cn.itcast.common.excel.model.ValueBean;

@ExcelHeader(headerName = "测试Bean")
@ExcelWarning(warningInfo = {"警告信息"})
public class BaseErrorUser {
	@ExcelColumn(columnName="ID")
	private ValueBean id ;
	@ExcelColumn(columnName = "姓名")
	private ValueBean username ;
	@ExcelColumn(columnName="地址")
	private ValueBean address ;

	public ValueBean getId() {
		return id;
	}

	public void setId(ValueBean id) {
		this.id = id;
	}

	public ValueBean getUsername() {
		return username;
	}

	public void setUsername(ValueBean username) {
		this.username = username;
	}

	public ValueBean getAddress() {
		return address;
	}

	public void setAddress(ValueBean address) {
		this.address = address;
	}
}
