package cn.itcast.common.excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
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
		} else if(ExcelType.XLSX.equals(excelType)) {
			return ExcelManager.createExcelManager().exportContainDataExcel_XLSX(results, clazz);
		} else {
			/*
			 * 导出Excel应对一定量大数据策略2
			 * 分页签Sheet导出海量数据
			 * 导出数据后及时刷新内存
			 * 
			 */
			return ExcelManager.createExcelManager().exportContainDataExcel_SXLSX(results, clazz) ;
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
	public static Workbook exportExcelData(List<?> appDatas, Class<?> clazz, ExcelType excelType, String sheetNames) {
		
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
		String[] sheetResult = new String[] {sheetNames} ;
		results.put("columnNames", list) ;
		results.put("appDatas", appDatas) ;
		results.put("sheetNames", sheetResult) ;
		results.put("isBigData", false) ;
		results.put("pageSize", 0) ;
		if(ExcelType.XLS.equals(excelType)) {
			return ExcelManager.createExcelManager().exportContainDataExcel_XLS(results, clazz) ;
		} else if(ExcelType.XLSX.equals(excelType)) {
			return ExcelManager.createExcelManager().exportContainDataExcel_XLSX(results, clazz);
		} else {
			/*
			 * 导出Excel应对一定量大数据策略2
			 * 分页签Sheet导出海量数据
			 * 导出数据后及时刷新内存
			 * 
			 */
			return ExcelManager.createExcelManager().exportContainDataExcel_SXLSX(results, clazz) ;
		}
	}
	
	/**
	 *
	 * importExcelData: Excel 模板严格按照生成的模板格式 获取列头信息 即获取了遍历的集合
	 * 每行应该具有的数据列数，必须强制满足条件 即：导入的数据行 列数必须与列头数保持一致 需要解析数据中是否有错误标记位，有则全部去掉
	 * 
	 * @param filePath
	 * @param excelType Excel类型
	 * @param clazz 携带注解的Bean字节码文件
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws ParseException
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public List<Object> importAllExcelData(String filePath, ExcelType excelType, Class<?> clazz) throws FileNotFoundException, IOException, SecurityException, NoSuchFieldException, InstantiationException, IllegalAccessException {
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		List<Object> list = new ArrayList<Object>() ;
		
		Iterator<Sheet> it = null ;
		if(excelType.equals(ExcelType.XLS)) {
			POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(filePath));
			it = excelManager.getHSSFWorkbook(fs).sheetIterator();
		} else if(excelType.equals(ExcelType.XLSX)){
			it = excelManager.getXSSFWorkbook(filePath).sheetIterator() ;
		} else {
			it = excelManager.getSXSSFWorkbook(filePath).sheetIterator() ;
		}
		
		while(it.hasNext()) {
			list.addAll(excelManager.importExcelData(clazz, it.next())) ;
		}
		return list ;
	}

	/**
	 *
	 * importExcelData: Excel 模板严格按照生成的模板格式 获取列头信息 即获取了遍历的集合
	 * 每行应该具有的数据列数，必须强制满足条件 即：导入的数据行 列数必须与列头数保持一致 需要解析数据中是否有错误标记位，有则全部去掉
	 * 
	 * @param filePath
	 * @param excelType Excel类型
	 * @param clazz 携带注解的Bean字节码文件
	 * @throws FileNotFoundException
	 * @throws IOException
	 * @throws ParseException
	 * @throws IllegalAccessException
	 * @throws InstantiationException
	 * @throws NoSuchFieldException
	 * @throws SecurityException
	 */
	public List<Object> importAllExcelData(InputStream in, ExcelType excelType, Class<?> clazz) throws FileNotFoundException, IOException, SecurityException, NoSuchFieldException, InstantiationException, IllegalAccessException {
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		List<Object> list = new ArrayList<Object>() ;
		
		Iterator<Sheet> it = null ;
		if(excelType.equals(ExcelType.XLS)) {
			POIFSFileSystem fs = new POIFSFileSystem(in);
			it = excelManager.getHSSFWorkbook(fs).sheetIterator();
		} else if(excelType.equals(ExcelType.XLSX)){
			it = excelManager.getXSSFWorkbook(in).sheetIterator() ;
		} else {
			it = excelManager.getSXSSFWorkbook(in).sheetIterator() ;
		}
		
		while(it.hasNext()) {
			list.addAll(excelManager.importExcelData(clazz, it.next())) ;
		}
		return list ;
	}

	/**
	 * 
	 *  importExcelData:(根据Sheet名称指定遍历). 
	 *  @return_type:List<Object>
	 *  @author zhangtian  
	 *  @param in
	 *  @param clazz
	 *  @param sheetNames
	 *  @return
	 *  @throws IOException
	 *  @throws SecurityException
	 *  @throws NoSuchFieldException
	 *  @throws InstantiationException
	 *  @throws IllegalAccessException
	 */
	public List<Object> importExcelData(InputStream in, Class<?> clazz, String excelType, String... sheetNames) throws IOException,
		SecurityException, NoSuchFieldException, InstantiationException, IllegalAccessException {
	
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		if(ArrayUtils.isEmpty(sheetNames)) {
			throw new RuntimeException("请指定页签Sheet名称...") ;
		}
		
		Workbook workbook = null ;
		if(excelType.equals(ExcelType.XLS)) {
			POIFSFileSystem fs = new POIFSFileSystem(in);
			workbook = excelManager.getHSSFWorkbook(fs);
		} else if(excelType.equals(ExcelType.XLSX)){
			workbook = excelManager.getXSSFWorkbook(in);
		} else {
			workbook = excelManager.getSXSSFWorkbook(in) ;
		}
		
		Sheet[] sheets = new Sheet[sheetNames.length] ;
		for(String sheetName : sheetNames) {
			int i = 0; 
			sheets[i] = workbook.getSheet(sheetName) ;
			i++ ;
		}
		
		return excelManager.importExcelData(clazz, sheets);
	}
	
	/**
	 * 
	 * importExcelDate: Excel 模板严格按照生成的模板格式 获取列头信息 即获取了遍历的集合
	 * 每行应该具有的数据列数，必须强制满足条件 即：导入的数据行 列数必须与列头数保持一致 需要解析数据中是否有错误标记位，有则全部去掉
	 *  importExcelData:(根据Sheet索引指定遍历). 
	 *  @return_type:List<Object>
	 *  @author zhangtian  
	 *  @param in
	 *  @param clazz
	 *  @param sheetNames
	 *  @return
	 *  @throws IOException
	 *  @throws SecurityException
	 *  @throws NoSuchFieldException
	 *  @throws InstantiationException
	 *  @throws IllegalAccessException
	 */
	public List<Object> importExcelData(InputStream in, Class<?> clazz, String excelType, int... sheetIndexes) throws IOException,
		SecurityException, NoSuchFieldException, InstantiationException, IllegalAccessException {
	
		ExcelManager excelManager = ExcelManager.createExcelManager() ;
		
		if(ArrayUtils.isEmpty(sheetIndexes)) {
			throw new RuntimeException("请指定页签Sheet索引...") ;
		}
		
		Workbook workbook = null ;
		if(excelType.equals(ExcelType.XLS)) {
			POIFSFileSystem fs = new POIFSFileSystem(in);
			workbook = excelManager.getHSSFWorkbook(fs);
		} else if(excelType.equals(ExcelType.XLSX)){
			workbook = excelManager.getXSSFWorkbook(in);
		} else {
			workbook = excelManager.getSXSSFWorkbook(in) ;
		}
		
		Sheet[] sheets = new Sheet[sheetIndexes.length] ;
		for(int sheetIndex : sheetIndexes) {
			int i = 0; 
			sheets[i] = workbook.getSheetAt(sheetIndex) ;
			i++ ;
		}
		
		return excelManager.importExcelData(clazz, sheets);
	}
}
