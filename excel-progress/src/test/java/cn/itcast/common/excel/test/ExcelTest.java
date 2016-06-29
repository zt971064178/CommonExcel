package cn.itcast.common.excel.test;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
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
		for(int i = 0; i < 10000; i++) {
			BaseUser u = new BaseUser() ;
			u.setId(UUID.randomUUID().toString());
			u.setUsername("Demo"+(i+1));
			u.setAddress("阳澄湖岛主"+i);
			appDatas.add(u) ;
		}
		
		
		Workbook workbook = ExcelUtils.exportExcelData(appDatas, BaseUser.class, ExcelType.XLS, true, 3000, new String[]{"test"}) ;
		OutputStream out = new FileOutputStream(new File("C:\\Users\\zhangtian\\Desktop\\demo.xls")) ;
		workbook.write(out);
		out.flush();
		out.close(); 
		workbook.close();
	}
}
