package cn.itcast.common.excel.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.UUID;

import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import cn.itcast.common.excel.ExcelUtils;
import cn.itcast.common.excel.constants.ExcelType;

/**
 * ClassName: ExcelTest  
 * (单元测试Excel导入导出)
 * @author zhangtian  
 * @version
 */
public class ExcelTest {
	
	/**
	 * 导出Excel应对一定量大数据策略1
	 * 分页签Sheet导出海量数据
	 * 问题：workbook中的数据流无法在内存中被清除
	 * 内存问题的该方法解决方案为：大数据时，分Excel导出，即导出多个超量数据Excel
	 */
	@Test
	public void testExportExcel() throws Exception {
		BaseUser u1 = new BaseUser() ;
		u1.setId(UUID.randomUUID().toString());
		u1.setUsername("张田");
		u1.setAddress("园区莲花五区");
		
		BaseUser u2 = new BaseUser() ;
		u2.setId(UUID.randomUUID().toString());
		u2.setUsername("小静静");
		u2.setAddress("崇明岛主");
		
		BaseUser u3 = new BaseUser() ;
		u3.setId(UUID.randomUUID().toString());
		u3.setUsername("王刚");
		u3.setAddress("阳澄湖岛主");
		
		List<BaseUser> appDatas = new ArrayList<BaseUser>() ;
		appDatas.add(u1) ;
		appDatas.add(u2) ;
		appDatas.add(u3) ;
		for(int i = 0; i < 100; i++) {
			BaseUser u = new BaseUser() ;
			u.setId(UUID.randomUUID().toString());
			u.setUsername("Demo"+(i+1));
			u.setAddress("阳澄湖岛主"+i);
			appDatas.add(u) ;
		}
		
		long startTime = new Date().getTime() ;
		Workbook workbook = ExcelUtils.exportExcelData(appDatas, BaseUser.class, ExcelType.OTHER, "zhangtian") ;
		OutputStream out = new FileOutputStream(new File("C:\\Users\\zhangtian\\Desktop\\demo.xlsx")) ;
		System.out.println(new Date().getTime() - startTime);
		workbook.write(out);
		out.flush();
		out.close(); 
		workbook.close();
	}
	
	/**
	 * 导出Excel应对一定量大数据策略2
	 * 分页签Sheet导出海量数据
	 * 导出数据后及时刷新内存
	 */
	@Test
	public void testExportExcelBigData() {
		
	}
	
}
