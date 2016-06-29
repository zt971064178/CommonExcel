package cn.itcast.common.excel.test;

import cn.itcast.common.excel.annotation.ExcelColumn;
import cn.itcast.common.excel.annotation.ExcelHeader;
import cn.itcast.common.excel.annotation.ExcelWarning;

/**
 * ClassName: BaseUser  
 * (测试Bean)
 * @author zhangtian  
 * @version
 */
@ExcelHeader(headerName = "测试Bean")
@ExcelWarning(warningInfo = {"警告信息"})
public class BaseUser {
	@ExcelColumn(columnName="ID")
	private String id ;
	@ExcelColumn(columnName = "姓名")
	private String username ;
	@ExcelColumn(columnName="地址")
	private String address ;

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getUsername() {
		return username;
	}

	public void setUsername(String username) {
		this.username = username;
	}

	public String getAddress() {
		return address;
	}

	public void setAddress(String address) {
		this.address = address;
	}
}
