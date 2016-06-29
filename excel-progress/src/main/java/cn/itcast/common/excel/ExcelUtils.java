package cn.itcast.common.excel;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.ss.usermodel.Workbook;

import cn.itcast.common.excel.annotation.ExcelColumn;
import cn.itcast.common.excel.annotation.ExcelHeader;
import cn.itcast.common.excel.annotation.ExcelWarning;
import cn.itcast.common.excel.constants.ExcelType;
import cn.itcast.common.excel.model.CellColumnValue;

/**
 * ClassName: ExcelUtils  
 * (Excel创建导入导出工具类)
 * @author zhangtian  
 * @version
 */
public class ExcelUtils {
	
	/**
	 * 
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param isBigData 是否开启大数据分页，true：是  false：否
	 * @param pageSize 分页每个页签显示的数据条数
	 * @return
	 * @author zhangtian
	 */
	public static Workbook exportExcelData(List<?> appDatas, Class<?> clazz, ExcelType excelType, boolean isBigData, int pageSize) {
		
		Map<String, Object> results = new HashMap<String, Object>() ;
		Field[] fields = clazz.getDeclaredFields() ;
		
		// 保存标题
		if(clazz.isAnnotationPresent(ExcelHeader.class)) {
			ExcelHeader excelHeader = clazz.getAnnotation(ExcelHeader.class) ;
			results.put("headerName", excelHeader.headerName()) ;
		}
		
		// 保存警告信息
		if(clazz.isAnnotationPresent(ExcelWarning.class)) {
			ExcelWarning excelWarning = clazz.getAnnotation(ExcelWarning.class) ;
			results.put("warningInfo", excelWarning.warningInfo()) ;
		}
		
		// 保存列头信息
		List<CellColumnValue> list = new ArrayList<CellColumnValue>() ;
		for(Field field : fields) {
			if(field.isAnnotationPresent(ExcelColumn.class)) {
				CellColumnValue cellColumnValue = new CellColumnValue() ;
				ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class) ;
				cellColumnValue.setColumnKey(field.getName());
				if(excelColumn.columnName() == null || "".equals(excelColumn.columnName().trim())) {
					cellColumnValue.setColumnValue(field.getName().toUpperCase());
				} else {
					cellColumnValue.setColumnValue(excelColumn.columnName());
				}
				list.add(cellColumnValue) ;
			}
		}
		
		// 大批量数据条件下的分割Sheet
		String[] sheetResult = null ; 
		if(isBigData) {
			int size = appDatas.size() ;
			int sheetNums = size % pageSize == 0 ? size / pageSize : (size / pageSize +1) ;
			sheetResult = new String[sheetNums] ;
			if(sheetNums > 1) {
				for(int i = 0; i< sheetNums; i++) {
					sheetResult[i] = ExcelManager.DEFAULT_SHEET_NAME + (i+1) ;
				}
			}
		} else {
			sheetResult = new String[]{ExcelManager.DEFAULT_SHEET_NAME+1} ;
		}
		
		results.put("columnNames", list) ;
		results.put("appDatas", appDatas) ;
		results.put("sheetNames", sheetResult) ;
		results.put("isBigData", isBigData) ;
		results.put("pageSize", pageSize) ;
		if(ExcelType.XLS.equals(excelType)) {
			return ExcelManager.createExcelManager().exportContainDataExcel_XLS(results, clazz) ;
		} else {
			return ExcelManager.createExcelManager().exportContainDataExcel_XLSX(results, clazz);
		}
	}
	
	/**
	 * 
	 * getExcelModalInfo:(获取Excel的头部标题以及列头信息)
	 *
	 * @param clazz 注解的Bean字节码
	 * @param appDatas 携带注解的Bean的数据集合
	 * @param excelType Excel文件类型XLS/XLSX
	 * @param sheetNames 自定义Sheet页签的名称
	 * @return
	 * @author zhangtian
	 */
	public static Workbook exportExcelData(List<?> appDatas, Class<?> clazz, ExcelType excelType, String[] sheetNames) {
		
		Map<String, Object> results = new HashMap<String, Object>() ;
		Field[] fields = clazz.getDeclaredFields() ;
		
		// 保存标题
		if(clazz.isAnnotationPresent(ExcelHeader.class)) {
			ExcelHeader excelHeader = clazz.getAnnotation(ExcelHeader.class) ;
			results.put("headerName", excelHeader.headerName()) ;
		}
		
		// 保存警告信息
		if(clazz.isAnnotationPresent(ExcelWarning.class)) {
			ExcelWarning excelWarning = clazz.getAnnotation(ExcelWarning.class) ;
			results.put("warningInfo", excelWarning.warningInfo()) ;
		}
		
		// 保存列头信息
		List<CellColumnValue> list = new ArrayList<CellColumnValue>() ;
		for(Field field : fields) {
			if(field.isAnnotationPresent(ExcelColumn.class)) {
				CellColumnValue cellColumnValue = new CellColumnValue() ;
				ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class) ;
				cellColumnValue.setColumnKey(field.getName());
				if(excelColumn.columnName() == null || "".equals(excelColumn.columnName().trim())) {
					cellColumnValue.setColumnValue(field.getName().toUpperCase());
				} else {
					cellColumnValue.setColumnValue(excelColumn.columnName());
				}
				list.add(cellColumnValue) ;
			}
		}
		
		// 大批量数据条件下的分割Sheet
		String[] sheetResult = sheetNames ;
		results.put("columnNames", list) ;
		results.put("appDatas", appDatas) ;
		results.put("sheetNames", sheetResult) ;
		if(ExcelType.XLS.equals(excelType)) {
			return ExcelManager.createExcelManager().exportContainDataExcel_XLS(results, clazz) ;
		} else {
			return ExcelManager.createExcelManager().exportContainDataExcel_XLSX(results, clazz);
		}
	}
}
