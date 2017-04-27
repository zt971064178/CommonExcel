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
@ExcelHeader(headerName = "测试小汽车")
@ExcelWarning(warningInfo = {"警告信息"})
public class BaseCar {
	@ExcelColumn(columnName="ID")
	private String id ;
	@ExcelColumn(columnName = "车品牌")
	private String carBand ;

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public String getCarBand() {
		return carBand;
	}

	public void setCarBand(String carBand) {
		this.carBand = carBand;
	}
}
